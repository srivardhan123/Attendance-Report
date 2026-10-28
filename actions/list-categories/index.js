const { Core } = require('@adobe/aio-sdk')
const { commerceFetch, commerceToken } = require('../lib/commerce')

// Lists merchandising categories (level > 1 skips "Root Catalog" and "Default Category").
async function main (params) {
  const logger = Core.Logger('list-categories', { level: params.LOG_LEVEL || 'info' })
  try {
    logger.info('Action invoked')
    const token = await commerceToken(params)
    const query = [
      'searchCriteria[filter_groups][0][filters][0][field]=level',
      'searchCriteria[filter_groups][0][filters][0][value]=1',
      'searchCriteria[filter_groups][0][filters][0][condition_type]=gt',
      'searchCriteria[pageSize]=500'
    ].join('&')
    const data = await commerceFetch(params, token, 'GET', `/categories/list?${query}`)
    const items = data.items || []

    // Label nested categories with their full path, e.g. "Men > Tops > Jackets".
    const byId = new Map(items.map(c => [String(c.id), c.name]))
    const categories = items
      .filter(c => c.is_active !== false)
      .map(c => {
        const ancestors = String(c.path || '').split('/').slice(2) // drop root + default category
        const label = ancestors.map(id => byId.get(id)).filter(Boolean).join(' > ') || c.name
        return { id: String(c.id), name: c.name, label }
      })
      .sort((a, b) => a.label.localeCompare(b.label))

    logger.info(`Returning ${categories.length} categories`)
    return { statusCode: 200, body: { categories } }
  } catch (error) {
    logger.error('Action failed:', error.message)
    return { statusCode: 500, body: { error: error.message } }
  }
}

exports.main = main
