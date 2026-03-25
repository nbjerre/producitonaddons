using Microsoft.AspNetCore.Mvc.Formatters;
using Microsoft.AspNetCore.Mvc;
using System;
using System.IO;
using System.Threading.Tasks;

namespace WebAPI.Configuration
{
    public class TextPlainInputFormatter : InputFormatter
    {
        public TextPlainInputFormatter()
        {
            SupportedMediaTypes.Add("text/plain");
        }

        protected override bool CanReadType(Type type) => true;

        public override async Task<InputFormatterResult> ReadRequestBodyAsync(InputFormatterContext context)
        {
            using var reader = new StreamReader(context.HttpContext.Request.Body);
            var body = await reader.ReadToEndAsync();
            try
            {
                var result = System.Text.Json.JsonSerializer.Deserialize(body, context.ModelType,
                    new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                return await InputFormatterResult.SuccessAsync(result);
            }
            catch
            {
                return await InputFormatterResult.FailureAsync();
            }
        }
    }
}