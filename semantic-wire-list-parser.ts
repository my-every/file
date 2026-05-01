function normalizeComparableValue(value: string | null | undefined): string {
    return String(value ?? '').trim().toUpperCase()
}

export function normalizeBrandingWireNo(wireNo: string | null | undefined, isNegative?: boolean): string {
    const normalized = String(wireNo ?? '').trim()
    if (!normalized) {
        return ''
    }

    if (normalized.startsWith('-')) {
        return normalized
    }

    if (isNegative || /^\d/.test(normalized)) {
        return `-${normalized}`
    }

    return normalized
}

export function normalizeBrandingGaugeSize(gaugeSize: string | null | undefined): string {
    return normalizeComparableValue(gaugeSize)
}

export function normalizeBrandingWireColor(wireColor: string | null | undefined): string {
    return normalizeComparableValue(wireColor)
}

export function buildBrandingRowGroupKey(input: {
    bundleName: string
    toLocation?: string | null
    gaugeSize?: string | null
    wireColor?: string | null
}): string {
    return [
        normalizeComparableValue(input.bundleName),
        normalizeComparableValue(input.toLocation),
        normalizeBrandingGaugeSize(input.gaugeSize),
        normalizeBrandingWireColor(input.wireColor),
    ].join('|')
}

export function buildBrandingRowGroupLabel(input: {
    bundleName: string
    toLocation?: string | null
    currentSheetName?: string | null
    gaugeSize?: string | null
    wireColor?: string | null
}): string {
    const bundleName = String(input.bundleName ?? '').trim()
    const toLocation = String(input.toLocation ?? '').trim()
    const currentSheetName = String(input.currentSheetName ?? '').trim()
    const gaugeSize = String(input.gaugeSize ?? '').trim()
    const wireColor = String(input.wireColor ?? '').trim()

    const baseLabel = toLocation && toLocation.toUpperCase() !== currentSheetName.toUpperCase()
        ? `${bundleName} - ${toLocation}`
        : bundleName

    const suffix = [gaugeSize, wireColor].filter(Boolean).join(' / ')
    return suffix ? `${baseLabel} - ${suffix}` : baseLabel
}

export function compareBrandingGaugeAndColor(
    left: { gaugeSize?: string | null; wireColor?: string | null },
    right: { gaugeSize?: string | null; wireColor?: string | null },
): number {
    const gaugeCompare = normalizeBrandingGaugeSize(left.gaugeSize).localeCompare(
        normalizeBrandingGaugeSize(right.gaugeSize),
        undefined,
        { numeric: true, sensitivity: 'base' },
    )
    if (gaugeCompare !== 0) {
        return gaugeCompare
    }

    return normalizeBrandingWireColor(left.wireColor).localeCompare(
        normalizeBrandingWireColor(right.wireColor),
        undefined,
        { numeric: true, sensitivity: 'base' },
    )
}
