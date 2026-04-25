using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace LitchiOzonRecovery
{
    public sealed class RussianImageAdaptationOptions
    {
        public string ApiBaseUrl { get; set; }
        public string ApiKey { get; set; }
        public string ResponsesModel { get; set; }
        public string ImageModel { get; set; }
        public string OutputDirectory { get; set; }
        public string Size { get; set; }
        public string Quality { get; set; }
        public string OutputFormat { get; set; }
        public int TimeoutMs { get; set; }

        public static RussianImageAdaptationOptions FromEnvironment()
        {
            return new RussianImageAdaptationOptions
            {
                ApiBaseUrl = FirstNonEmpty(
                    Environment.GetEnvironmentVariable("CODEXMANAGER_BASE_URL"),
                    Environment.GetEnvironmentVariable("OPENAI_BASE_URL"),
                    "http://127.0.0.1:48760"),
                ApiKey = FirstNonEmpty(
                    Environment.GetEnvironmentVariable("CODEXMANAGER_API_KEY"),
                    Environment.GetEnvironmentVariable("OPENAI_API_KEY")),
                ResponsesModel = FirstNonEmpty(
                    Environment.GetEnvironmentVariable("RUSSIAN_IMAGE_RESPONSES_MODEL"),
                    "gpt-5.4-mini"),
                ImageModel = FirstNonEmpty(
                    Environment.GetEnvironmentVariable("RUSSIAN_IMAGE_MODEL"),
                    "gpt-image-1.5"),
                OutputDirectory = FirstNonEmpty(
                    Environment.GetEnvironmentVariable("RUSSIAN_IMAGE_OUTPUT_DIR"),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "generated-images")),
                Size = FirstNonEmpty(Environment.GetEnvironmentVariable("RUSSIAN_IMAGE_SIZE"), "1024x1024"),
                Quality = FirstNonEmpty(Environment.GetEnvironmentVariable("RUSSIAN_IMAGE_QUALITY"), "medium"),
                OutputFormat = FirstNonEmpty(Environment.GetEnvironmentVariable("RUSSIAN_IMAGE_OUTPUT_FORMAT"), "jpeg"),
                TimeoutMs = 180000
            };
        }

        private static string FirstNonEmpty(params string[] values)
        {
            for (int i = 0; i < values.Length; i++)
            {
                if (!string.IsNullOrWhiteSpace(values[i]))
                {
                    return values[i].Trim();
                }
            }

            return string.Empty;
        }
    }

    public sealed class RussianImageAdaptationRequest
    {
        public string ProductTitle { get; set; }
        public string RussianTitle { get; set; }
        public string Category { get; set; }
        public string SourceImageUrl { get; set; }
        public string SourceImagePath { get; set; }
        public string OutputFileNamePrefix { get; set; }
    }

    public sealed class RussianImageAdaptationResult
    {
        public bool Success { get; set; }
        public string ImagePath { get; set; }
        public string ImageUrl { get; set; }
        public string Base64Image { get; set; }
        public string RevisedPrompt { get; set; }
        public string RawResponse { get; set; }
        public string ErrorMessage { get; set; }
    }

    public sealed class RussianImageAdaptationModule
    {
        public RussianImageAdaptationResult AdaptProductImage(
            RussianImageAdaptationRequest request,
            RussianImageAdaptationOptions options)
        {
            RussianImageAdaptationResult result = new RussianImageAdaptationResult();
            RussianImageAdaptationOptions resolved = options ?? RussianImageAdaptationOptions.FromEnvironment();

            if (request == null)
            {
                result.ErrorMessage = "Russian image adaptation request is required.";
                return result;
            }

            if (string.IsNullOrWhiteSpace(resolved.ApiKey))
            {
                result.ErrorMessage = "CODEXMANAGER_API_KEY or OPENAI_API_KEY is required.";
                return result;
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(request.SourceImagePath) && File.Exists(request.SourceImagePath))
                {
                    return EditLocalImage(request, resolved);
                }

                return GenerateOrEditWithResponses(request, resolved);
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                return result;
            }
        }

        private static RussianImageAdaptationResult GenerateOrEditWithResponses(
            RussianImageAdaptationRequest request,
            RussianImageAdaptationOptions options)
        {
            JObject payload = new JObject();
            payload["model"] = options.ResponsesModel;
            payload["stream"] = false;

            JArray tools = new JArray();
            JObject imageTool = new JObject();
            imageTool["type"] = "image_generation";
            imageTool["model"] = options.ImageModel;
            imageTool["size"] = options.Size;
            imageTool["quality"] = options.Quality;
            imageTool["output_format"] = options.OutputFormat;
            imageTool["moderation"] = "auto";
            tools.Add(imageTool);
            payload["tools"] = tools;

            JArray content = new JArray();
            content.Add(new JObject(
                new JProperty("type", "input_text"),
                new JProperty("text", BuildPrompt(request))));

            if (!string.IsNullOrWhiteSpace(request.SourceImageUrl))
            {
                content.Add(new JObject(
                    new JProperty("type", "input_image"),
                    new JProperty("image_url", request.SourceImageUrl)));
            }

            JArray input = new JArray();
            input.Add(new JObject(
                new JProperty("role", "user"),
                new JProperty("content", content)));
            payload["input"] = input;

            string response = PostJson(BuildUrl(options.ApiBaseUrl, "/v1/responses"), payload.ToString(Formatting.None), options);
            return ParseImageResponse(response, request, options);
        }

        private static RussianImageAdaptationResult EditLocalImage(
            RussianImageAdaptationRequest request,
            RussianImageAdaptationOptions options)
        {
            Dictionary<string, string> fields = new Dictionary<string, string>();
            fields["model"] = options.ImageModel;
            fields["prompt"] = BuildPrompt(request);
            fields["size"] = options.Size;
            fields["quality"] = options.Quality;
            fields["output_format"] = options.OutputFormat;
            fields["moderation"] = "auto";

            string response = PostMultipart(
                BuildUrl(options.ApiBaseUrl, "/v1/images/edits"),
                fields,
                "image",
                request.SourceImagePath,
                options);
            return ParseImageResponse(response, request, options);
        }

        private static string BuildPrompt(RussianImageAdaptationRequest request)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("Create an Ozon marketplace product photo adapted for Russian consumers. ");
            builder.Append("Keep the same product identity, shape, color, material, and key visible features. ");
            builder.Append("Make the scene feel natural for Russia: clean apartment or marketplace lifestyle context, Cyrillic-safe visual style, winter/neutral daylight when useful, no Chinese marketplace watermarks, no 1688 UI, no fake brand logos, no certification marks, no misleading packaging claims. ");
            builder.Append("Use realistic commercial photography, product centered, no text unless already present on the product, no extra accessories that change the sold item. ");
            builder.Append("Product title: ").Append(request.ProductTitle ?? string.Empty).Append(". ");

            if (!string.IsNullOrWhiteSpace(request.RussianTitle))
            {
                builder.Append("Russian listing title: ").Append(request.RussianTitle).Append(". ");
            }

            if (!string.IsNullOrWhiteSpace(request.Category))
            {
                builder.Append("Category: ").Append(request.Category).Append(". ");
            }

            if (!string.IsNullOrWhiteSpace(request.SourceImageUrl) || !string.IsNullOrWhiteSpace(request.SourceImagePath))
            {
                builder.Append("Edit the input image rather than inventing a different product.");
            }
            else
            {
                builder.Append("Generate a new product photo based on the product details.");
            }

            return builder.ToString();
        }

        private static RussianImageAdaptationResult ParseImageResponse(
            string response,
            RussianImageAdaptationRequest request,
            RussianImageAdaptationOptions options)
        {
            RussianImageAdaptationResult result = new RussianImageAdaptationResult();
            result.RawResponse = response;

            JObject root = JObject.Parse(response);
            string imageUrl = Convert.ToString(root.SelectToken("data[0].url") ?? string.Empty);
            string base64 = Convert.ToString(root.SelectToken("data[0].b64_json") ?? string.Empty);
            string revisedPrompt = Convert.ToString(root.SelectToken("data[0].revised_prompt") ?? string.Empty);

            if (string.IsNullOrEmpty(base64))
            {
                JArray output = root["output"] as JArray;
                for (int i = 0; output != null && i < output.Count; i++)
                {
                    JObject item = output[i] as JObject;
                    if (item == null)
                    {
                        continue;
                    }

                    if (string.Equals(Convert.ToString(item["type"]), "image_generation_call", StringComparison.OrdinalIgnoreCase))
                    {
                        base64 = Convert.ToString(item["result"] ?? string.Empty);
                        revisedPrompt = Convert.ToString(item["revised_prompt"] ?? revisedPrompt);
                        break;
                    }
                }
            }

            result.ImageUrl = imageUrl;
            result.Base64Image = base64;
            result.RevisedPrompt = revisedPrompt;

            if (!string.IsNullOrEmpty(base64))
            {
                result.ImagePath = SaveBase64Image(base64, request, options);
                result.Success = true;
                return result;
            }

            if (!string.IsNullOrEmpty(imageUrl))
            {
                result.Success = true;
                return result;
            }

            result.ErrorMessage = "Image response did not include b64_json, image_generation_call.result, or data[0].url.";
            return result;
        }

        private static string SaveBase64Image(
            string base64,
            RussianImageAdaptationRequest request,
            RussianImageAdaptationOptions options)
        {
            if (!Directory.Exists(options.OutputDirectory))
            {
                Directory.CreateDirectory(options.OutputDirectory);
            }

            string extension = "." + (string.IsNullOrWhiteSpace(options.OutputFormat) ? "png" : options.OutputFormat.Trim().ToLowerInvariant());
            string prefix = string.IsNullOrWhiteSpace(request.OutputFileNamePrefix)
                ? "russian-image"
                : SanitizeFileName(request.OutputFileNamePrefix);
            string path = Path.Combine(options.OutputDirectory, prefix + "-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + extension);
            File.WriteAllBytes(path, Convert.FromBase64String(StripDataUrlPrefix(base64)));
            return path;
        }

        private static string StripDataUrlPrefix(string value)
        {
            int comma = value == null ? -1 : value.IndexOf(',');
            if (comma > 0 && value.Substring(0, comma).IndexOf("base64", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return value.Substring(comma + 1);
            }

            return value ?? string.Empty;
        }

        private static string SanitizeFileName(string value)
        {
            char[] invalid = Path.GetInvalidFileNameChars();
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < value.Length; i++)
            {
                char c = value[i];
                builder.Append(Array.IndexOf(invalid, c) >= 0 ? '_' : c);
            }

            return builder.Length == 0 ? "russian-image" : builder.ToString();
        }

        private static string PostJson(string url, string json, RussianImageAdaptationOptions options)
        {
            HttpWebRequest request = CreateRequest(url, options);
            request.Method = "POST";
            request.ContentType = "application/json";

            byte[] body = Encoding.UTF8.GetBytes(json);
            request.ContentLength = body.Length;
            using (Stream stream = request.GetRequestStream())
            {
                stream.Write(body, 0, body.Length);
            }

            return ReadResponse(request);
        }

        private static string PostMultipart(
            string url,
            IDictionary<string, string> fields,
            string fileFieldName,
            string filePath,
            RussianImageAdaptationOptions options)
        {
            string boundary = "----OzonPilotBoundary" + Guid.NewGuid().ToString("N");
            HttpWebRequest request = CreateRequest(url, options);
            request.Method = "POST";
            request.ContentType = "multipart/form-data; boundary=" + boundary;

            using (Stream stream = request.GetRequestStream())
            {
                foreach (KeyValuePair<string, string> pair in fields)
                {
                    WriteString(stream, "--" + boundary + "\r\n");
                    WriteString(stream, "Content-Disposition: form-data; name=\"" + pair.Key + "\"\r\n\r\n");
                    WriteString(stream, pair.Value ?? string.Empty);
                    WriteString(stream, "\r\n");
                }

                WriteString(stream, "--" + boundary + "\r\n");
                WriteString(stream, "Content-Disposition: form-data; name=\"" + fileFieldName + "\"; filename=\"" + Path.GetFileName(filePath) + "\"\r\n");
                WriteString(stream, "Content-Type: " + DetectImageContentType(filePath) + "\r\n\r\n");
                byte[] fileBytes = File.ReadAllBytes(filePath);
                stream.Write(fileBytes, 0, fileBytes.Length);
                WriteString(stream, "\r\n--" + boundary + "--\r\n");
            }

            return ReadResponse(request);
        }

        private static HttpWebRequest CreateRequest(string url, RussianImageAdaptationOptions options)
        {
            EnsureModernTls();
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Accept = "application/json";
            request.Timeout = options.TimeoutMs <= 0 ? 180000 : options.TimeoutMs;
            request.Headers["Authorization"] = "Bearer " + options.ApiKey;
            return request;
        }

        private static void WriteString(Stream stream, string value)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(value);
            stream.Write(bytes, 0, bytes.Length);
        }

        private static string DetectImageContentType(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".jpg" || ext == ".jpeg")
            {
                return "image/jpeg";
            }

            if (ext == ".webp")
            {
                return "image/webp";
            }

            return "image/png";
        }

        private static string BuildUrl(string baseUrl, string path)
        {
            string cleanBase = string.IsNullOrWhiteSpace(baseUrl) ? "http://127.0.0.1:48760" : baseUrl.TrimEnd('/');
            return cleanBase + path;
        }

        private static string ReadResponse(HttpWebRequest request)
        {
            try
            {
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                using (Stream stream = response.GetResponseStream())
                using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (WebException ex)
            {
                string detail = ex.Message;
                if (ex.Response != null)
                {
                    using (Stream stream = ex.Response.GetResponseStream())
                    using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        detail = reader.ReadToEnd();
                    }
                }

                throw new InvalidOperationException(detail, ex);
            }
        }

        private static void EnsureModernTls()
        {
            const int tls10 = 192;
            const int tls11 = 768;
            const int tls12 = 3072;
            ServicePointManager.SecurityProtocol =
                (SecurityProtocolType)(tls10 | tls11 | tls12);
            ServicePointManager.Expect100Continue = false;
        }
    }
}
