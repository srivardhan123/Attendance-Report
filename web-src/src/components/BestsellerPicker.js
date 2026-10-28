import React from 'react'
import {
  View, Flex, Heading, Text, Image, Badge, ToggleButton, ProgressCircle,
  IllustratedMessage, Content, InlineAlert
} from '@adobe/react-spectrum'
import NotFound from '@spectrum-icons/illustrations/NotFound'
import { formatPrice } from './api'

const BASIS_TEXT = {
  recent: (days) => `Ranked by units sold in the last ${days} days.`,
  'all-time': (days) => `No sales in the last ${days} days — ranked by all-time units sold.`,
  newest: () => 'No sales recorded in this category yet — showing the newest products as style references.'
}

export default function BestsellerPicker ({ categoryLabel, loading, error, result, selectedSku, onSelect }) {
  if (!categoryLabel) {
    return (
      <IllustratedMessage>
        <NotFound />
        <Heading>Pick a category</Heading>
        <Content>We'll pull recent sales and show what's selling best there.</Content>
      </IllustratedMessage>
    )
  }
  if (loading) {
    return (
      <Flex direction="column" alignItems="center" justifyContent="center" height="size-3000" gap="size-150">
        <ProgressCircle aria-label="Loading best-sellers" isIndeterminate size="L" />
        <Text>Analyzing recent orders in {categoryLabel}…</Text>
      </Flex>
    )
  }
  if (error) {
    return (
      <InlineAlert variant="negative">
        <Heading>Couldn't load best-sellers</Heading>
        <Content>{error}</Content>
      </InlineAlert>
    )
  }
  if (!result) return null
  if (result.bestsellers.length === 0) {
    return (
      <IllustratedMessage>
        <NotFound />
        <Heading>No products with photos in {categoryLabel}</Heading>
        <Content>Choose another category to find a style reference.</Content>
      </IllustratedMessage>
    )
  }

  return (
    <Flex direction="column" gap="size-150">
      <Text>{(BASIS_TEXT[result.basis] || BASIS_TEXT.recent)(result.days)}</Text>
      <Flex wrap gap="size-200">
        {result.bestsellers.map((p, i) => {
          const selected = p.sku === selectedSku
          return (
            <View
              key={p.sku}
              width="size-3000"
              borderWidth={selected ? 'thick' : 'thin'}
              borderColor={selected ? 'focus' : 'mid'}
              borderRadius="medium"
              padding="size-150"
              backgroundColor="gray-50"
            >
              <Flex direction="column" gap="size-100">
                <View height="size-2400" borderRadius="small" overflow="hidden" backgroundColor="gray-75">
                  <Image src={p.imageUrl} alt={p.name} objectFit="contain" height="100%" />
                </View>
                <Flex alignItems="center" gap="size-100">
                  <Badge variant={i === 0 ? 'positive' : 'neutral'}>#{i + 1}</Badge>
                  <Text UNSAFE_style={{ fontWeight: 700 }}>{p.name}</Text>
                </Flex>
                <Text>
                  {formatPrice(p.price)}
                  {result.basis !== 'newest' ? ` · ${p.unitsSold} sold · ${formatPrice(p.revenue)}` : ''}
                </Text>
                <ToggleButton
                  isSelected={selected}
                  onChange={() => onSelect(p.sku)}
                  isEmphasized
                  aria-label={`Style after ${p.name}`}
                >
                  {selected ? 'Styling after this' : 'Style after this'}
                </ToggleButton>
              </Flex>
            </View>
          )
        })}
      </Flex>
    </Flex>
  )
}
