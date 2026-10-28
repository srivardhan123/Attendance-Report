const { Core } = require('@adobe/aio-sdk')
const { commerceFetch, commerceToken } = require('../lib/commerce')

const IMAGE_TYPES = ['image/png', 'image/jpeg', 'image/gif']

function makeSku (name) {
  const slug = String(name).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 40)
  return `${slug || 'product'}-${Date.now().toString(36)}`
}

// Publishes the approved idea as a simple, in-stock, visible product with the generated hero image.
async function main (params) {
  const logger = Core.Logger('publish-product', { level: params.LOG_LEVEL || 'info' })
  try {
    logger.info('Action invoked')
    logger.debug('Params:', JSON.stringify({ name: params.name, price: params.price, categoryId: params.categoryId }))

    const missing = ['name', 'price', 'categoryId', 'imageUrl'].filter(p => params[p] === undefined || params[p] === '')
    if (missing.length > 0) {
      return { statusCode: 400, body: { error: `Missing required params: ${missing.join(', ')}` } }
    }
    const price = Number(params.price)
    if (!Number.isFinite(price) || price <= 0) {
      return { statusCode: 400, body: { error: 'price must be a positive number' } }
    }
    const qty = Number(params.qty) > 0 ? Math.floor(Number(params.qty)) : 100

    // Download the approved image first so a stale URL fails before anything is created.
    const imgRes = await fetch(params.imageUrl)
    if (!imgRes.ok) throw new Error(`Could not download approved image (${imgRes.status}) — try regenerating`)
    let imageType = (imgRes.headers.get('content-type') || 'image/png').split(';')[0].trim()
    if (!IMAGE_TYPES.includes(imageType)) imageType = 'image/png'
    const base64 = Buffer.from(await imgRes.arrayBuffer()).toString('base64')

    const token = await commerceToken(params)
    const sku = params.sku || makeSku(params.name)
    const description = params.description ||
      (params.styledAfter ? `${params.name} — styled after our best-seller ${params.styledAfter}.` : params.name)

    const product = await commerceFetch(params, token, 'POST', '/products', {
      product: {
        sku,
        name: params.name,
        attribute_set_id: 4,
        price,
        status: 1,
        visibility: 4,
        type_id: 'simple',
        weight: 1,
        extension_attributes: {
          website_ids: [1],
          category_links: [{ position: 0, category_id: String(params.categoryId) }],
          stock_item: { qty, is_in_stock: true, manage_stock: true }
        },
        custom_attributes: [
          { attribute_code: 'url_key', value: sku },
          { attribute_code: 'description', value: description }
        ]
      },
      saveOptions: true
    })
    logger.info(`Created product ${sku} (id ${product.id})`)

    let imageAttached = true
    let warning
    try {
      await commerceFetch(params, token, 'POST', `/products/${encodeURIComponent(sku)}/media`, {
        entry: {
          media_type: 'image',
          label: params.name,
          position: 1,
          disabled: false,
          types: ['image', 'small_image', 'thumbnail'],
          content: { base64_encoded_data: base64, type: imageType, name: `${sku}.${imageType.split('/')[1]}` }
        }
      })
    } catch (e) {
      imageAttached = false
      warning = `Product created, but attaching the image failed: ${e.message}`
      logger.error(warning)
    }

    logger.info('Action completed successfully')
    return {
      statusCode: 200,
      body: { id: product.id, sku, name: params.name, price, qty, categoryId: String(params.categoryId), imageAttached, warning }
    }
  } catch (error) {
    logger.error('Action failed:', error.message)
    return { statusCode: error.statusCode === 400 ? 400 : 500, body: { error: error.message } }
  }
}

exports.main = main
