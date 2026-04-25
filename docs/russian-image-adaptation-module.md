# Russian image adaptation module

This module generates or edits product photos for Russian/Ozon marketplace context through the local CodexManager OpenAI-compatible API.

## Files

- `src/LitchiOzonRecovery/RussianImageAdaptationModule.cs`
- `src/LitchiOzonRecovery/ProductImagePublisher.cs`
- `src/LitchiOzonRecovery/LitchiOzonRecovery.csproj`
- `deploy/ozon-image-server/install-ubuntu.sh`
- `deploy/ozon-image-server/configure-existing-nginx-port80.sh`

## Required environment

CodexManager must be running locally:

```powershell
$env:CODEXMANAGER_BASE_URL = "http://127.0.0.1:48760"
$env:CODEXMANAGER_API_KEY = "<local codexmanager api key>"
```

Optional settings:

```powershell
$env:RUSSIAN_IMAGE_RESPONSES_MODEL = "gpt-5.4-mini"
$env:RUSSIAN_IMAGE_MODEL = "gpt-image-1.5"
$env:RUSSIAN_IMAGE_SIZE = "1024x1024"
$env:RUSSIAN_IMAGE_QUALITY = "medium"
$env:RUSSIAN_IMAGE_OUTPUT_FORMAT = "jpeg"
$env:RUSSIAN_IMAGE_OUTPUT_DIR = "D:\ozon-pilot-images"
```

Image hosting settings for the public Ubuntu server:

```powershell
$env:OZON_IMAGE_UPLOAD_URL = "http://47.76.248.181/ozon-upload"
$env:OZON_IMAGE_PUBLIC_BASE_URL = "http://47.76.248.181/ozon-images/"
$env:OZON_IMAGE_UPLOAD_TOKEN = "<same token configured on server>"
```

If `CODEXMANAGER_API_KEY` is not available, the module falls back to `OPENAI_API_KEY`. If `CODEXMANAGER_BASE_URL` is not set, it defaults to `http://127.0.0.1:48760`.

## Basic usage

```csharp
RussianImageAdaptationModule module = new RussianImageAdaptationModule();

RussianImageAdaptationResult image = module.AdaptProductImage(
    new RussianImageAdaptationRequest
    {
        ProductTitle = product.Title,
        RussianTitle = product.RussianTitle,
        Category = product.Keyword,
        SourceImageUrl = product.MainImage,
        OutputFileNamePrefix = product.OfferId
    },
    RussianImageAdaptationOptions.FromEnvironment());

if (image.Success)
{
    // image.ImageUrl: direct remote URL if the API returned one.
    // image.ImagePath: local file when the API returned base64.
    // image.Base64Image: raw image payload for custom upload/storage.
}
```

To publish a generated local file to the public Ubuntu server:

```csharp
IProductImagePublisher publisher = new HttpProductImagePublisher(
    HttpProductImagePublisherOptions.FromEnvironment());

if (image.Success && !string.IsNullOrEmpty(image.ImagePath))
{
    string publicUrl = publisher.Publish(image.ImagePath);
}
```

## Editing a local image

Use `SourceImagePath` when a downloaded source image exists locally:

```csharp
RussianImageAdaptationResult image = module.AdaptProductImage(
    new RussianImageAdaptationRequest
    {
        ProductTitle = product.Title,
        RussianTitle = product.RussianTitle,
        Category = product.Keyword,
        SourceImagePath = @"D:\source-images\1688-123.png",
        OutputFileNamePrefix = product.OfferId
    },
    RussianImageAdaptationOptions.FromEnvironment());
```

Local edits call `/v1/images/edits`. URL-based adaptation calls `/v1/responses` with the `image_generation` tool.

## Hook point for Ozon upload

Current Ozon payload is built in `ProductAutomationService.BuildOzonImportItem`. The clean integration point is before calling `BuildOzonImportItem`, inside `UploadToOzon`, after text enrichment:

```csharp
EnrichProductWithDeepSeek(product, options, categoryAttributes);

RussianImageAdaptationResult adapted = imageModule.AdaptProductImage(
    new RussianImageAdaptationRequest
    {
        ProductTitle = product.Title,
        RussianTitle = product.RussianTitle,
        Category = product.Keyword,
        SourceImageUrl = product.MainImage,
        OutputFileNamePrefix = product.OfferId
    },
    RussianImageAdaptationOptions.FromEnvironment());

if (adapted.Success && !string.IsNullOrEmpty(adapted.ImageUrl))
{
    product.MainImage = adapted.ImageUrl;
    if (!product.Images.Contains(adapted.ImageUrl))
    {
        product.Images.Insert(0, adapted.ImageUrl);
    }
}
```

If the image API returns only a local file, publish it before replacing Ozon images:

```csharp
string publicImageUrl = adapted.ImageUrl;
if (adapted.Success && string.IsNullOrEmpty(publicImageUrl) && !string.IsNullOrEmpty(adapted.ImagePath))
{
    IProductImagePublisher publisher = new HttpProductImagePublisher(
        HttpProductImagePublisherOptions.FromEnvironment());
    publicImageUrl = publisher.Publish(adapted.ImagePath);
}

if (!string.IsNullOrEmpty(publicImageUrl))
{
    product.MainImage = publicImageUrl;
    if (!product.Images.Contains(publicImageUrl))
    {
        product.Images.Insert(0, publicImageUrl);
    }
}
```

Important: Ozon import expects image URLs reachable by Ozon. The configured public Ubuntu server is intended to provide those URLs.

## Ubuntu image server setup

Run on the server `47.76.248.181` as root or a sudo-capable user:

```bash
UPLOAD_TOKEN='replace-with-long-random-token' bash deploy/ozon-image-server/install-ubuntu.sh
```

This repository has also been configured for the existing port-80 Nginx site on `47.76.248.181`:

- health: `http://47.76.248.181/ozon-image-health`
- upload: `http://47.76.248.181/ozon-upload`
- public images: `http://47.76.248.181/ozon-images/`

Verify from local machine:

```powershell
Invoke-WebRequest -UseBasicParsing http://47.76.248.181/ozon-image-health
```

If the base service is already installed and only the existing port-80 Nginx site needs to be patched:

```bash
UPLOAD_TOKEN='same-token' bash deploy/ozon-image-server/configure-existing-nginx-port80.sh
```

## What the prompt enforces

The module asks the image model to:

- keep product identity, shape, color, material, and visible features
- adapt the setting for Russian consumers and Ozon listing use
- remove Chinese marketplace watermarks and 1688 UI context
- avoid fake logos, certifications, package claims, and misleading accessories
- use realistic commercial product photography

## PR/module delivery note

Keep this module isolated in a PR:

- add `RussianImageAdaptationModule.cs`
- add `ProductImagePublisher.cs`
- add the `.csproj` compile include
- add this document
- keep production upload wiring as an explicit follow-up so listing owners can choose when generated images replace original supplier images
