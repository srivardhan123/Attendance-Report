const { Core } = require('@adobe/aio-sdk')
const { ffToken, ffUpload, ffPoll } = require('../lib/firefly')

const SUPPORTED_TYPES = ['image/jpeg', 'image/png', 'image/webp', 'image/tiff']

function buildPrompt ({ productName, category, styleNotes }) {
  return [
    `Professional e-commerce product photograph of ${productName}, a new ${category} product.`,
    'Single product, centered, fully in frame, soft even studio lighting,',
    'plain seamless light background, crisp detail, catalog hero shot.',
    styleNotes ? `Style notes: ${styleNotes}.` : ''
  ].join(' ').trim()
}

// Generates a product image for the new idea, styled after the chosen best-seller's photo.
async function main (params) {
  const logger = Core.Logger('generate-image', { level: params.LOG_LEVEL || 'info' })
  try {
    logger.info('Action invoked')
    logger.debug('Params:', JSON.stringify({ productName: params.productName, category: params.category, referenceImageUrl: params.referenceImageUrl }))

    const missing = ['productName', 'category', 'referenceImageUrl'].filter(p => !params[p])
    if (missing.length > 0) {
      return { statusCode: 400, body: { error: `Missing required params: ${missing.join(', ')}` } }
    }

    // Fetch the best-seller photo and upload it to Firefly storage (uploadId avoids the URL allowlist).
    const refRes = await fetch(params.referenceImageUrl)
    if (!refRes.ok) throw new Error(`Could not download best-seller image (${refRes.status})`)
    const contentType = (refRes.headers.get('content-type') || 'image/jpeg').split(';')[0].trim()
    if (!SUPPORTED_TYPES.includes(contentType)) {
      return { statusCode: 400, body: { error: `Best-seller image type ${contentType} is not supported as a style reference` } }
    }
    const bytes = Buffer.from(await refRes.arrayBuffer())

    const token = await ffToken(params)
    const clientId = params.PHOTOSHOP_CLIENT_ID
    const uploadId = await ffUpload(token, clientId, bytes, contentType)

    const seed = Number.isInteger(Number(params.seed)) && params.seed !== undefined && params.seed !== ''
      ? Number(params.seed)
      : Math.floor(Math.random() * 1000000)
    const prompt = buildPrompt(params)
    const job = await ffPoll(token, clientId, '/v3/images/generate-async', {
      prompt,
      numVariations: 1,
      seeds: [seed],
      size: { width: 1024, height: 1024 },
      contentClass: 'photo',
      style: { imageReference: { source: { uploadId } }, strength: 60 }
    })

    const output = job.result && job.result.outputs && job.result.outputs[0]
    if (!output || !output.image || !output.image.url) {
      throw new Error('Firefly returned no image')
    }

    logger.info('Action completed successfully')
    return { statusCode: 200, body: { imageUrl: output.image.url, seed: output.seed ?? seed, prompt } }
  } catch (error) {
    logger.error('Action failed:', error.message)
    return { statusCode: 500, body: { error: error.message } }
  }
}

exports.main = main
