import React from 'react'
import { View, Flex, Heading, Text, Image, Button, ButtonGroup, StatusLight, Divider } from '@adobe/react-spectrum'
import { formatPrice } from './api'

export default function ReviewPanel ({ brief, reference, image, isPublishing, onApprove, onRegenerate, onBack }) {
  return (
    <Flex direction="column" gap="size-300">
      <Flex wrap gap="size-400" alignItems="start">
        <View flex="1 1 380px">
          <Heading level={3} marginTop={0}>Your new product</Heading>
          <View UNSAFE_className="checkerboard" borderWidth="thin" borderColor="mid" borderRadius="medium" height="size-5000">
            <Image src={image.finalUrl} alt={`Generated hero image for ${brief.name}`} objectFit="contain" height="100%" />
          </View>
          <Flex direction="column" marginTop="size-150" gap="size-50">
            <Text UNSAFE_style={{ fontWeight: 700 }}>{brief.name}</Text>
            <Text>{brief.categoryLabel} · {formatPrice(brief.price)}</Text>
            <StatusLight variant="positive">Color corrected · background removed</StatusLight>
          </Flex>
        </View>
        <View flex="1 1 380px">
          <Heading level={3} marginTop={0}>Styled after best-seller</Heading>
          <View borderWidth="thin" borderColor="mid" borderRadius="medium" height="size-5000" backgroundColor="gray-75">
            <Image src={reference.imageUrl} alt={reference.name} objectFit="contain" height="100%" />
          </View>
          <Flex direction="column" marginTop="size-150" gap="size-50">
            <Text UNSAFE_style={{ fontWeight: 700 }}>{reference.name}</Text>
            <Text>
              SKU {reference.sku} · {formatPrice(reference.price)}
              {reference.unitsSold > 0 ? ` · ${reference.unitsSold} sold` : ''}
            </Text>
          </Flex>
        </View>
      </Flex>
      <Divider size="S" />
      <ButtonGroup>
        <Button variant="accent" onPress={onApprove} isPending={isPublishing}>Approve & publish</Button>
        <Button variant="secondary" onPress={onRegenerate} isDisabled={isPublishing}>Regenerate</Button>
        <Button variant="secondary" isQuiet onPress={onBack} isDisabled={isPublishing}>Edit brief</Button>
      </ButtonGroup>
    </Flex>
  )
}
