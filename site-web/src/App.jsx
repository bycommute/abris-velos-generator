import React, { useState } from 'react'

function App() {
  const [file, setFile] = useState(null)
  const [fileName, setFileName] = useState('nepastoucher.xlsx')
  const [isGenerating, setIsGenerating] = useState(false)
  const [status, setStatus] = useState(null)
  const [progress, setProgress] = useState(0)
  const [downloadUrl, setDownloadUrl] = useState(null)

  const handleFileUpload = (event) => {
    const uploadedFile = event.target.files[0]
    if (uploadedFile && uploadedFile.name.endsWith('.xlsx')) {
      setFile(uploadedFile)
      setFileName(uploadedFile.name)
      setStatus({ type: 'success', message: `Fichier "${uploadedFile.name}" prêt à être utilisé` })
    } else {
      setStatus({ type: 'error', message: 'Veuillez sélectionner un fichier Excel (.xlsx)' })
    }
  }

  const handleGenerate = async () => {
    setIsGenerating(true)
    setStatus({ type: 'loading', message: 'Génération en cours...' })
    setProgress(0)
    setDownloadUrl(null)

    try {
      // Créer FormData pour envoyer le fichier
      const formData = new FormData()
      if (file) {
        formData.append('file', file)
      }

      // Appeler l'API (Netlify Function ou API séparée)
      // Pour développement local, utiliser: http://localhost:5000/generate
      const apiUrl = import.meta.env.VITE_API_URL || '/.netlify/functions/generate'
      const response = await fetch(apiUrl, {
        method: 'POST',
        body: formData
      })

      if (!response.ok) {
        throw new Error('Erreur lors de la génération')
      }

      // Récupérer le blob du ZIP
      const blob = await response.blob()
      const url = window.URL.createObjectURL(blob)
      setDownloadUrl(url)
      setStatus({ type: 'success', message: 'Génération terminée ! Vous pouvez télécharger le fichier ZIP.' })
      setProgress(100)
    } catch (error) {
      setStatus({ type: 'error', message: `Erreur: ${error.message}` })
      setProgress(0)
    } finally {
      setIsGenerating(false)
    }
  }

  const handleDownload = () => {
    if (downloadUrl) {
      const a = document.createElement('a')
      a.href = downloadUrl
      a.download = 'abris-velos-variantes.zip'
      document.body.appendChild(a)
      a.click()
      document.body.removeChild(a)
    }
  }

  return (
    <div className="app">
      <h1>🏗️ Générateur d'Abris Vélos</h1>
      <p className="subtitle">Générez toutes les variantes d'abris vélos à partir d'un fichier Excel de base</p>

      <div className="section">
        <h2>📁 Fichier de base</h2>
        <div className={`file-info ${file ? 'has-file' : ''}`}>
          <div>
            <div className="file-name">{fileName}</div>
            {file && <div className="file-size">{(file.size / 1024).toFixed(2)} KB</div>}
          </div>
          <label>
            <input
              type="file"
              accept=".xlsx"
              onChange={handleFileUpload}
              disabled={isGenerating}
            />
            <button className="upload-btn" disabled={isGenerating}>
              {file ? 'Changer le fichier' : 'Uploader un fichier'}
            </button>
          </label>
        </div>
        {!file && (
          <p style={{ color: '#666', fontSize: '0.9em', marginTop: '10px' }}>
            Le fichier par défaut "nepastoucher.xlsx" sera utilisé si aucun fichier n'est uploadé.
          </p>
        )}
      </div>

      <div className="section">
        <h2>⚙️ Génération</h2>
        <button
          className="generate-btn"
          onClick={handleGenerate}
          disabled={isGenerating}
        >
          {isGenerating ? 'Génération en cours...' : '🚀 Générer toutes les variantes'}
        </button>

        {status && (
          <div className={`status ${status.type}`}>
            {status.message}
          </div>
        )}

        {isGenerating && (
          <div className="progress">
            <div className="progress-bar" style={{ width: `${progress}%` }}></div>
          </div>
        )}

        {downloadUrl && (
          <button
            className="download-btn"
            onClick={handleDownload}
          >
            📥 Télécharger le fichier ZIP
          </button>
        )}
      </div>
    </div>
  )
}

export default App

