import React, { useEffect, useMemo, useRef, useState } from 'react'
import {
  Provider, defaultTheme, View, Flex, Heading, Text, Content, Form, TextField, NumberField,
  ComboBox, Item, Button, ButtonGroup, InlineAlert, Image, ProgressBar, Divider, Well, IllustratedMessage
} from '@adobe/react-spectrum'
import actions from '../config.json'
import { invokeAction, formatPrice } from './api'
import BestsellerPicker from './BestsellerPicker'
import ReviewPanel from './ReviewPanel'

const STEPS = ['Generating image styled after the best-seller', 'Correcting color & lighting, removing background']

export default function App ({ runtime, ims }) {
  // Do NOT call runtime.done() here — index.js calls it in the ready handler
  const urls = {
    categories: actions['list-categories'],
    bestsellers: actions['get-bestsellers'],
    generate: actions['generate-image'],
    enhance: actions['enhance-image'],
    publish: actions['publish-product']
  }
  const backendMissing = Object.values(urls).some(u => !u)

  const [name, setName] = useState('')
  const [price, setPrice] = useState(NaN)
  const [categoryId, setCategoryId] = useState(null)

  const [categories, setCategories] = useState([])
  const [categoriesLoading, setCategoriesLoading] = useState(false)
  const [categoriesError, setCategoriesError] = useState(null)

  const [bestsellers, setBestsellers] = useState(null)
  const [bestsellersLoading, setBestsellersLoading] = useState(false)
  const [bestsellersError, setBestsellersError] = useState(null)
  const [selectedSku, setSelectedSku] = useState(null)
  const bestsellerRequest = useRef(0)

  // stage: brief -> generating -> review -> published
  const [stage, setStage] = useState('brief')
  const [step, setStep] = useState(0)
  const [image, setImage] = useState(null)
  const [genError, setGenError] = useState(null)
  const [isPublishing, setIsPublishing] = useState(false)
  const [publishError, setPublishError] = useState(null)
  const [published, setPublished] = useState(null)

  const category = categories.find(c => c.id === categoryId) || null
  const reference = useMemo(
    () => (bestsellers && bestsellers.bestsellers.find(p => p.sku === selectedSku)) || null,
    [bestsellers, selectedSku]
  )

  useEffect(() => {
    if (!urls.categories) return
    setCategoriesLoading(true)
    invokeAction(urls.categories, ims, {})
      .then(data => setCategories((data && data.categories) || []))
      .catch(e => setCategoriesError(e.message))
      .finally(() => setCategoriesLoading(false))
  }, [urls.categories])

  async function loadBestsellers (id) {
    const requestId = ++bestsellerRequest.current
    setBestsellers(null)
    setSelectedSku(null)
    setBestsellersError(null)
    if (!id || !urls.bestsellers) return
    setBestsellersLoading(true)
    try {
      const data = await invokeAction(urls.bestsellers, ims, { categoryId: id, days: 90, limit: 4 })
      if (requestId !== bestsellerRequest.current) return
      setBestsellers(data)
      // Default to the top seller; the merchandiser can switch.
      setSelectedSku(data.bestsellers[0] ? data.bestsellers[0].sku : null)
    } catch (e) {
      if (requestId === bestsellerRequest.current) setBestsellersError(e.message)
    } finally {
      if (requestId === bestsellerRequest.current) setBestsellersLoading(false)
    }
  }

  function onCategoryChange (key) {
    const id = key == null ? null : String(key)
    if (id === categoryId) return
    setCategoryId(id)
    loadBestsellers(id)
  }

  async function generate () {
    setStage('generating')
    setGenError(null)
    setPublishError(null)
    setStep(0)
    try {
      const gen = await invokeAction(urls.generate, ims, {
        productName: name.trim(),
        category: category.name,
        referenceImageUrl: reference.imageUrl
      })
      setStep(1)
      const clean = await invokeAction(urls.enhance, ims, { imageUrl: gen.imageUrl })
      setImage({ rawUrl: gen.imageUrl, finalUrl: clean.finalUrl, seed: gen.seed })
      setStage('review')
    } catch (e) {
      setGenError(e.message)
      setStage(image ? 'review' : 'brief')
    }
  }

  async function approve () {
    setIsPublishing(true)
    setPublishError(null)
    try {
      const result = await invokeAction(urls.publish, ims, {
        name: name.trim(),
        price,
        categoryId,
        imageUrl: image.finalUrl,
        qty: 100,
        styledAfter: reference.name
      })
      setPublished(result)
      setStage('published')
    } catch (e) {
      setPublishError(e.message)
    } finally {
      setIsPublishing(false)
    }
  }

  function startOver () {
    setName('')
    setPrice(NaN)
    setCategoryId(null)
    setBestsellers(null)
    setSelectedSku(null)
    setImage(null)
    setPublished(null)
    setGenError(null)
    setPublishError(null)
    setStage('brief')
  }

  const briefComplete = name.trim() && category && Number(price) > 0 && reference
  const canGenerate = briefComplete && urls.generate && urls.enhance
  const brief = { name: name.trim(), price, categoryLabel: category ? category.label : '' }

  return (
    <Provider theme={defaultTheme} colorScheme="light">
      <View padding="size-400" maxWidth="1200px" marginX="auto">
        <Heading level={1} marginBottom="size-50">Merchandising Copilot</Heading>
        <Text>Describe a new product idea, see what's selling in its category, and launch it styled after a proven winner.</Text>

        {backendMissing && (
          <View marginTop="size-300">
            <InlineAlert variant="info">
              <Heading>Backend actions are not live yet</Heading>
              <Content>To use backend actions in the live preview, select a workspace using 'Set deploy target' below, then click 'Test App'. If a workspace is already selected, click 'Test App' to start the backend.</Content>
            </InlineAlert>
          </View>
        )}

        {stage === 'published' && published && (
          <View marginTop="size-400">
            <InlineAlert variant="positive">
              <Heading>{published.name} is live in the catalog</Heading>
              <Content>
                SKU {published.sku} · {formatPrice(published.price)} · {published.qty} in stock · enabled and visible in catalog & search.
                {published.warning ? ` ${published.warning}` : ''}
              </Content>
            </InlineAlert>
            <Flex wrap gap="size-300" marginTop="size-300">
              <View UNSAFE_className="checkerboard" width="size-3600" height="size-3600" borderRadius="medium" borderWidth="thin" borderColor="mid">
                <Image src={image.finalUrl} alt={published.name} objectFit="contain" height="100%" />
              </View>
            </Flex>
            <ButtonGroup marginTop="size-300">
              <Button variant="accent" onPress={startOver}>Launch another product</Button>
            </ButtonGroup>
          </View>
        )}

        {stage === 'review' && image && reference && (
          <View marginTop="size-400">
            <Heading level={2}>Preview</Heading>
            {genError && (
              <View marginBottom="size-200">
                <InlineAlert variant="negative"><Heading>Regeneration failed</Heading><Content>{genError}</Content></InlineAlert>
              </View>
            )}
            {publishError && (
              <View marginBottom="size-200">
                <InlineAlert variant="negative"><Heading>Publishing failed</Heading><Content>{publishError}</Content></InlineAlert>
              </View>
            )}
            <ReviewPanel
              brief={brief}
              reference={reference}
              image={image}
              isPublishing={isPublishing}
              onApprove={approve}
              onRegenerate={generate}
              onBack={() => { setStage('brief'); setGenError(null) }}
            />
          </View>
        )}

        {stage === 'generating' && (
          <View marginTop="size-400">
            <Well>
              <Flex direction="column" gap="size-200">
                <Heading level={2} margin={0}>Creating {name.trim() || 'your product'}</Heading>
                <ProgressBar label={STEPS[step]} isIndeterminate width="100%" />
                <Text>Step {step + 1} of {STEPS.length} · this usually takes under a minute.</Text>
              </Flex>
            </Well>
          </View>
        )}

        {stage === 'brief' && (
          <Flex wrap gap="size-500" marginTop="size-400" alignItems="start">
            <View flex="1 1 300px" maxWidth="size-4600">
              <Heading level={2} marginTop={0}>1. Product idea</Heading>
              <Form
                aria-label="New product idea"
                onSubmit={(e) => { e.preventDefault(); if (canGenerate) generate() }}
                isRequired
              >
                <TextField label="Product name" value={name} onChange={setName} placeholder="e.g. Coastal Linen Shirt" />
                <ComboBox
                  label="Category"
                  items={categories}
                  selectedKey={categoryId}
                  onSelectionChange={onCategoryChange}
                  loadingState={categoriesLoading ? 'loading' : 'idle'}
                  isDisabled={!urls.categories}
                  placeholder="Start typing a category"
                  validationState={categoriesError ? 'invalid' : undefined}
                  errorMessage={categoriesError ? `Couldn't load categories: ${categoriesError}` : undefined}
                >
                  {item => <Item key={item.id} textValue={item.label}>{item.label}</Item>}
                </ComboBox>
                <NumberField
                  label="Price"
                  value={price}
                  onChange={setPrice}
                  minValue={0.01}
                  formatOptions={{ style: 'currency', currency: 'USD' }}
                />
                <ButtonGroup marginTop="size-200">
                  <Button variant="accent" type="submit" isDisabled={!canGenerate}>
                    Generate product image
                  </Button>
                </ButtonGroup>
                {!briefComplete && (
                  <Text>Fill in all three fields and pick a best-seller to style after.</Text>
                )}
              </Form>
              {genError && (
                <View marginTop="size-200">
                  <InlineAlert variant="negative"><Heading>Image generation failed</Heading><Content>{genError}</Content></InlineAlert>
                </View>
              )}
            </View>

            <Divider orientation="vertical" size="S" />

            <View flex="2 1 480px">
              <Heading level={2} marginTop={0}>
                2. What's selling well{category ? ` in ${category.label}` : ''}
              </Heading>
              {!urls.bestsellers && category
                ? (
                  <IllustratedMessage>
                    <Heading>Sales data needs the backend</Heading>
                    <Content>Start the backend to see best-sellers for this category.</Content>
                  </IllustratedMessage>
                  )
                : (
                  <BestsellerPicker
                    categoryLabel={category ? category.label : ''}
                    loading={bestsellersLoading}
                    error={bestsellersError}
                    result={bestsellers}
                    selectedSku={selectedSku}
                    onSelect={setSelectedSku}
                  />
                  )}
            </View>
          </Flex>
        )}
      </View>
    </Provider>
  )
}
