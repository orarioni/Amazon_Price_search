@{
    MarketplaceId     = 'A1VC38T7YXB528'
    SpApiBaseUrl      = 'https://sellingpartnerapi-fe.amazon.com' #real world
    #SpApiBaseUrl      = 'https://sandbox.sellingpartnerapi-fe.amazon.com' #sandbox
    UserAgent         = 'AmazonPriceTool/0.4'
    MaxRetries        = 4
    RetryMaxAttempts   = 6
    RetryMaxWaitSec    = 120
    CatalogBatchSize  = 20
    PricingMinIntervalSec = 2.2
    DebugSpApiResponse = $false
    DebugSpApiResponseMaxChars = 4000
    JanAsinCacheTtlHours = 168
    OfferCacheTtlHours   = 24
    NegativeCacheTtlHours = 12
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
