using Microsoft.AspNetCore.Razor.TagHelpers;

namespace Nop.Web.Framework.TagHelpers
{
    [HtmlTargetElement("label", Attributes = PostfixAttributeName)]
    public class PostfixTagHelper : TagHelper
    {
        private const string PostfixAttributeName = "asp-postfix";

        [HtmlAttributeName(PostfixAttributeName)]
        public string Postfix { get; set; }

        public override void Process(TagHelperContext context, TagHelperOutput output)
        {
            base.Process(context, output);

            output.Content.Append(Postfix);
        }

        public override int Order
        {
            get
            {
                return 100; // Needs to run after aspnet
            }
        }
    }
}