if (typeof jQuery !== "undefined" && typeof saveAs !== "undefined") {
	(function ($) {
		$.fn.wordExport = function (fileName) {
			fileName =
				typeof fileName !== "undefined" ? fileName : "Test-Engine-Export";
			var static = {
				mhtml: {
					top:
						"Mime-Version: 1.0\nContent-Base: " +
						location.href +
						'\nContent-Type: Multipart/related; boundary="NEXT.ITEM-BOUNDARY";type="text/html"\n\n--NEXT.ITEM-BOUNDARY\nContent-Type: text/html; charset="utf-8"\nContent-Location: ' +
						location.href +
						"\n\n<!DOCTYPE html>\n<html>\n_html_</html>",
					head: '<head>\n<meta http-equiv="Content-Type" content="text/html; charset=utf-8">\n<style>\n_styles_\n</style>\n</head>\n',
					body: "<body>_body_</body>",
				},
			};
			var options = {
				maxWidth: 624,
			};

			// Clone selected element before manipulating it
			var markup = $(this).clone();

			// Remove hidden elements from the output
			markup.each(function () {
				var self = $(this);
				if (self.is(":hidden")) self.remove();
			});

			// Simple approach: resize images without converting to base64
			var img = markup.find("img");
			for (var i = 0; i < img.length; i++) {
				// Just resize images, keep original src
				var w = Math.min(img[i].width || options.maxWidth, options.maxWidth);
				var h = img[i].height
					? img[i].height * (w / (img[i].width || w))
					: "auto";

				$(img[i]).css({
					"max-width": w + "px",
					height: h === "auto" ? "auto" : h + "px",
				});
			}

			// No image data to embed - images will reference original URLs
			var mhtmlBottom = "\n--NEXT.ITEM-BOUNDARY--";

			// Basic CSS for better Word compatibility
			var styles = `
                body { font-family: Arial, sans-serif; }
                img { max-width: 100%; height: auto; }
                table { border-collapse: collapse; width: 100%; }
                td, th { border: 1px solid #ddd; padding: 8px; }
            `;

			// Aggregate parts of the file together
			var fileContent =
				static.mhtml.top.replace(
					"_html_",
					static.mhtml.head.replace("_styles_", styles) +
						static.mhtml.body.replace("_body_", markup.html())
				) + mhtmlBottom;

			// Create a Blob with the file contents
			var blob = new Blob([fileContent], {
				type: "application/msword;charset=utf-8",
			});
			saveAs(blob, fileName + ".doc");
		};
	})(jQuery);
} else {
	if (typeof jQuery === "undefined") {
		console.error("Test Paper: missing dependency (jQuery)");
	}
	if (typeof saveAs === "undefined") {
		console.error("Test Paper Export: missing dependency (FileSaver.js)");
	}
}
