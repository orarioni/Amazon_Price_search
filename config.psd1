@{
    MarketplaceId     = 'A1VC38T7YXB528'
    SpApiBaseUrl      = 'https://sellingpartnerapi-fe.amazon.com'
    UserAgent         = 'AmazonPriceTool/0.4'
    MaxRetries        = 4
    CatalogBatchSize  = 20
    PricingBatchSize  = 20
    PricingSingleFallbackThreshold = 3
    CacheTtlHours     = 24

    Paths = @{
        SecretsFile = 'secrets/lwa_secrets.xml'
        DataDir     = 'data'
        InputFile   = 'data/input.xlsx'
        OutputFile  = 'data/output.xlsx'
        LogDir      = 'logs'
        LogFile     = 'logs/run.log'
        CacheDir    = 'cache'
        CacheFile   = 'cache/price_cache.json'
        HistoryDir  = 'cache/history'
        AccessTokenCacheFile = 'cache/access_token.json'
    }
}
