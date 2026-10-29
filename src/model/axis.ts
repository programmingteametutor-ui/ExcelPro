/**
 * Axis metrics for a virtualized grid: maps row/column indices to pixel offsets and back
 * with sparse size overrides. Lookups are O(log n) in the number of overrides, so a sheet
 * with a million rows costs nothing until rows are resized or hidden.
 */
export class AxisMetrics {
  private keys: number[] = [] // sorted indices with non-default size
  private cum: number[] = [] // cum[i] = sum of (size - default) for keys[0..i-1]
  private sizes = new Map<number, number>()

  constructor(
    public defaultSize: number,
    public count: number,
    overrides: Record<number, number> = {},
    hidden: number[] = [],
  ) {
    for (const [k, v] of Object.entries(overrides)) this.sizes.set(+k, v)
    for (const h of hidden) this.sizes.set(h, 0)
    this.rebuild()
  }

  private rebuild() {
    this.keys = [...this.sizes.keys()].sort((a, b) => a - b)
    this.cum = new Array(this.keys.length + 1)
    this.cum[0] = 0
    for (let i = 0; i < this.keys.length; i++) {
      this.cum[i + 1] = this.cum[i] + (this.sizes.get(this.keys[i])! - this.defaultSize)
    }
  }

  sizeOf(i: number): number {
    return this.sizes.get(i) ?? this.defaultSize
  }

  /** Number of override keys strictly less than i. */
  private lowerBound(i: number): number {
    let lo = 0
    let hi = this.keys.length
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (this.keys[mid] < i) lo = mid + 1
      else hi = mid
    }
    return lo
  }

  offsetOf(i: number): number {
    return i * this.defaultSize + this.cum[this.lowerBound(i)]
  }

  get total(): number {
    return this.offsetOf(this.count)
  }

  /** Index of the item containing pixel offset px (clamped to [0, count-1]). */
  indexAt(px: number): number {
    if (px <= 0) return 0
    let lo = 0
    let hi = this.count - 1
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1
      if (this.offsetOf(mid) <= px) lo = mid
      else hi = mid - 1
    }
    // Skip zero-size (hidden) items forward
    while (lo < this.count - 1 && this.sizeOf(lo) === 0) lo++
    return lo
  }

  /** Next visible index in a direction, or the same index if none. */
  step(i: number, dir: 1 | -1): number {
    let j = i + dir
    while (j >= 0 && j < this.count && this.sizeOf(j) === 0) j += dir
    return j < 0 || j >= this.count ? i : j
  }
}
