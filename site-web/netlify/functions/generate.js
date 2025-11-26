const { exec } = require('child_process')
const { promisify } = require('util')
const fs = require('fs').promises
const path = require('path')
const archiver = require('archiver')

const execAsync = promisify(exec)

exports.handler = async (event, context) => {
  // Vérifier que c'est une requête POST
  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      body: JSON.stringify({ error: 'Method not allowed' })
    }
  }

  try {
    // Créer un dossier temporaire pour le travail
    const tempDir = `/tmp/generate-${Date.now()}`
    await fs.mkdir(tempDir, { recursive: true })

    // Copier les scripts Python depuis le projet parent
    const scriptsDir = path.join(__dirname, '../../../')
    const scripts = [
      'generate_bosquet_ouvert.py',
      'generate_bosquet_ferme.py',
      'generate_bosquet_ouvert_compact.py',
      'generate_bosquet_ferme_compact.py',
      'generate_domino_ouvert.py',
      'generate_domino_ferme.py',
      'generate_domino_ouvert_compact.py',
      'generate_domino_ferme_compact.py',
      'generate_metallique_ouvert.py',
      'generate_metallique_ferme.py',
      'generate_metallique_ouvert_compact.py',
      'generate_metallique_ferme_compact.py',
      'generate_neve_ouvert.py',
      'generate_neve_ferme.py',
      'generate_neve_ferme_compact.py'
    ]

    // Copier les scripts
    for (const script of scripts) {
      const src = path.join(scriptsDir, script)
      const dest = path.join(tempDir, script)
      try {
        await fs.copyFile(src, dest)
      } catch (err) {
        console.warn(`Script ${script} non trouvé, ignoré`)
      }
    }

    // Gérer le fichier uploadé
    let baseFile = path.join(tempDir, 'nepastoucher.xlsx')
    if (event.body && event.isBase64Encoded) {
      const fileBuffer = Buffer.from(event.body, 'base64')
      await fs.writeFile(baseFile, fileBuffer)
    } else if (event.body) {
      // Si le fichier est dans FormData
      const formData = event.body
      // Parser le FormData (simplifié)
      // En production, utiliser un parser FormData approprié
    }

    // Créer le dossier résultats
    const resultatsDir = path.join(tempDir, 'résultats')
    await fs.mkdir(resultatsDir, { recursive: true })

    // Créer le dossier fichier de base
    const baseDir = path.join(tempDir, 'fichier de base')
    await fs.mkdir(baseDir, { recursive: true })
    
    // Copier le fichier de base
    await fs.copyFile(baseFile, path.join(baseDir, 'nepastoucher.xlsx'))

    // Exécuter tous les scripts Python
    const pythonScripts = scripts.filter(s => s.endsWith('.py'))
    for (const script of pythonScripts) {
      try {
        const scriptPath = path.join(tempDir, script)
        await execAsync(`cd ${tempDir} && python3 ${scriptPath}`, {
          env: { ...process.env, PYTHONPATH: tempDir }
        })
      } catch (err) {
        console.error(`Erreur avec ${script}:`, err.message)
        // Continuer avec les autres scripts
      }
    }

    // Créer le fichier ZIP avec tous les résultats
    const zipPath = path.join(tempDir, 'resultats.zip')
    const output = require('fs').createWriteStream(zipPath)
    const archive = archiver('zip', { zlib: { level: 9 } })

    return new Promise((resolve, reject) => {
      archive.pipe(output)

      // Ajouter tous les fichiers Excel générés
      archive.directory(resultatsDir, false)

      archive.finalize()

      output.on('close', async () => {
        // Lire le fichier ZIP
        const zipBuffer = await fs.readFile(zipPath)

        // Nettoyer
        await fs.rm(tempDir, { recursive: true, force: true })

        return resolve({
          statusCode: 200,
          headers: {
            'Content-Type': 'application/zip',
            'Content-Disposition': 'attachment; filename="abris-velos-variantes.zip"'
          },
          body: zipBuffer.toString('base64'),
          isBase64Encoded: true
        })
      })

      archive.on('error', (err) => {
        reject({
          statusCode: 500,
          body: JSON.stringify({ error: `Erreur lors de la création du ZIP: ${err.message}` })
        })
      })
    })
  } catch (error) {
    console.error('Erreur:', error)
    return {
      statusCode: 500,
      body: JSON.stringify({ error: error.message })
    }
  }
}

