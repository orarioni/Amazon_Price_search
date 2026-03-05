# Function Dependencies

`scripts/lib/AmazonPriceLib.psm1` の全関数を、依存レイヤー順に整理した一覧です。

## 1. Entry point / Orchestration

- `Invoke-AmazonPriceUpdate`
  - depends on:
    - 認証: `Get-LwaAccessTokenCached`, `ConvertTo-PlainText`
    - Catalog: `Get-AsinMapByJanBatch`
    - Pricing: `Get-PriceMapByAsinSequential`（互換名: `Get-PriceMapByAsinBatch`）
    - Cache: `Import-PersistentCache`, `Save-PersistentCache`, `Get-JanCacheKey`, `Get-OfferCacheKey`, `Test-CacheFreshByStatus`
    - Output/History: `Add-DailyPriceHistory`, `Write-Log`

## 2. Catalog / Pricing domain

- `Get-AsinMapByJanBatch`
  - depends on: `Invoke-SpApiRequest`, `New-SpApiHeaders`, `Split-IntoChunks`, `Get-ErrorDetail`, `Get-StatusClassification`, `Get-LwaAccessTokenCached`
- `Get-PriceMapByAsinSequential`
  - depends on: `Get-PriceBySingleAsin`, `Get-ErrorDetail`
- `Get-PriceMapByAsinBatch`（互換ラッパー）
  - depends on: `Get-PriceMapByAsinSequential`
- `Get-PriceBySingleAsin`
  - depends on: `Wait-ForPricingSlot`, `Invoke-SpApiRequest`, `New-SpApiHeaders`, `Get-LowestNewPriceFromOffers`, `Get-ErrorDetail`, `Get-LwaAccessTokenCached`
- `Get-LowestNewPriceFromOffers`
  - depends on: (module内依存なし)

## 3. API / Retry / Error classification

- `Invoke-SpApiRequest`
  - depends on: `Get-ErrorDetail`, `Get-HeaderValue`, `Write-SpApiResponseDebugLog`, `Update-PricingThrottleFromLimit`, `Write-Log`
- `Invoke-WithRetry`
  - depends on: `Write-Log`
- `Get-ErrorDetail`
  - depends on: `Get-StatusClassification`
- `Get-StatusClassification`
  - depends on: (module内依存なし)

## 4. Auth / Headers / Time

- `Get-LwaAccessTokenCached`
  - depends on: `Read-AccessTokenCache`, `Save-AccessTokenCache`, `Get-LwaAccessToken`
- `Get-LwaAccessToken`
  - depends on: `Invoke-WithRetry`
- `Get-AmzDateHeaderValue`
  - depends on: (module内依存なし)
- `New-SpApiHeaders`
  - depends on: `Get-AmzDateHeaderValue`

## 5. Rate-limit throttling metrics

- `Initialize-RunStats`
- `Add-WaitMetric`
- `Wait-ForPricingSlot`
  - depends on: `Add-WaitMetric`
- `Update-PricingThrottleFromLimit`

## 6. Cache / History

- `Import-PersistentCache`
  - depends on: `Write-Log`
- `Save-PersistentCache`
  - depends on: `Write-Log`
- `Add-DailyPriceHistory`
  - depends on: `Write-Log`
- `Get-JanCacheKey`
- `Get-OfferCacheKey`
- `Get-CacheTtlHoursByStatus`
- `Test-CacheFreshByStatus`
  - depends on: `Get-CacheTtlHoursByStatus`, `Test-CacheFresh`
- `Test-CacheFresh`

## 7. Utility

- `ConvertTo-PlainText`
- `Write-Log`
- `Get-HeaderValue`
- `Get-PropertyValue`
- `Mask-SensitiveText`
- `Write-SpApiResponseDebugLog`
  - depends on: `Get-PropertyValue`, `Write-Log`, `Mask-SensitiveText`
- `Split-IntoChunks`
- `Read-AccessTokenCache`
- `Save-AccessTokenCache`
- `Save-SecretsInteractive`
- `Get-FunctionDependencyMap`
  - depends on: (module内依存なし。依存一覧データを返す)

## Dependency order (critical path)

1. `Invoke-AmazonPriceUpdate`
2. `Get-AsinMapByJanBatch`
3. `Get-PriceMapByAsinSequential`
4. `Get-PriceBySingleAsin`
5. `Invoke-SpApiRequest`
6. cache/history persist (`Save-PersistentCache`, `Add-DailyPriceHistory`)
