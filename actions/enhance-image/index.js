const { Core } = require('@adobe/aio-sdk')
const { psToken, psPoll, psOutputUrl } = require('../lib/photoshop')

// Cleans up a generated image: auto color/light correction, then background removal.
async function main (params) {
  const logger = Core.Logger('enhance-image', { level: params.LOG_LEVEL || 'info' })
  try {
    logger.info('Action invoked')

    const missing = ['imageUrl'].filter(p => !params[p])
    if (missing.length > 0) {
      return { statusCode: 400, body: { error: `Missing required params: ${missing.join(', ')}` } }
    }

    const token = await psToken(params)
    const clientId = params.PHOTOSHOP_CLIENT_ID

    const toneJob = await psPoll(token, clientId, 'https://photoshop-api.adobe.io/v2/edit', {
      image: { source: { url: params.imageUrl } },
      edits: { autoTone: true },
      outputs: [{ mediaType: 'image/png', destination: { validityPeriod: 86400 } }]
    })
    const correctedUrl = psOutputUrl(toneJob)
    logger.info('Color/light correction done')

    // remove-background accepts the /v2/edit short-url output directly.
    const cutoutJob = await psPoll(token, clientId, 'https://image.adobe.io/v2/remove-background', {
      image: { source: { url: correctedUrl } },
      mode: 'cutout',
      trim: false,
      output: { mediaType: 'image/png' }
    })
    const finalUrl = psOutputUrl(cutoutJob)

    logger.info('Action completed successfully')
    return { statusCode: 200, body: { correctedUrl, finalUrl } }
  } catch (error) {
    logger.error('Action failed:', error.message)
    return { statusCode: 500, body: { error: error.message } }
  }
}

exports.main = main
