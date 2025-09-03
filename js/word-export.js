exportWordBtn = document.querySelectorAll(".export-word");

exportWordBtn.forEach((element) => {
	element.addEventListener("click", () => {
		exportToDocx(document.getElementById("previewHTML"), "document");
	});
});

// Solution 1: Using FileReader API for local files
function convertImageToBase64(imgElement) {
	return new Promise((resolve, reject) => {
		// Check if it's a file:// URL
		if (imgElement.src.startsWith("file://")) {
			// For file protocol, we need to use FileReader
			convertLocalImageToBase64(imgElement.src).then(resolve).catch(reject);
		} else {
			// For HTTP/HTTPS URLs, use canvas method
			const canvas = document.createElement("canvas");
			const ctx = canvas.getContext("2d");
			const img = new Image();
			img.crossOrigin = "anonymous";

			img.onload = function () {
				canvas.width = img.width;
				canvas.height = img.height;
				ctx.drawImage(img, 0, 0);

				try {
					const base64 = canvas.toDataURL("image/png");
					resolve(base64);
				} catch (error) {
					reject(error);
				}
			};

			img.onerror = function () {
				reject(new Error("Failed to load image"));
			};

			img.src = imgElement.src.startsWith("http")
				? imgElement.src
				: new URL(imgElement.src, window.location.href).href;
		}
	});
}

// Helper function to convert local file to base64
function convertLocalImageToBase64(filePath) {
	return new Promise((resolve, reject) => {
		// Extract filename from file path
		const filename = filePath.split("/").pop().split("\\").pop();

		// Create file input to simulate file selection
		const input = document.createElement("input");
		input.type = "file";
		input.accept = "image/*";

		input.onchange = function (event) {
			const file = event.target.files[0];
			if (file) {
				const reader = new FileReader();
				reader.onload = function (e) {
					resolve(e.target.result);
				};
				reader.onerror = function () {
					reject(new Error("Failed to read file"));
				};
				reader.readAsDataURL(file);
			} else {
				reject(new Error("No file selected"));
			}
		};

		// This won't work automatically, but provides a fallback
		reject(
			new Error(
				`Cannot access local file: ${filename}. Please use HTTP server or modify image handling.`
			)
		);
	});
}

// Solution 2: Modified version that handles file protocol gracefully
function convertImageToBase64Fallback(imgElement) {
	return new Promise((resolve, reject) => {
		// If it's a file:// URL, try to read it as a relative path
		if (imgElement.src.startsWith("file://")) {
			// Convert file:// URL to relative path
			const relativePath = imgElement.src
				.replace("file://", "")
				.replace(
					window.location.pathname.substring(
						0,
						window.location.pathname.lastIndexOf("/")
					),
					""
				);

			// Try to load as relative path
			const img = new Image();
			img.crossOrigin = "anonymous";

			img.onload = function () {
				const canvas = document.createElement("canvas");
				const ctx = canvas.getContext("2d");
				canvas.width = img.width;
				canvas.height = img.height;
				ctx.drawImage(img, 0, 0);

				try {
					const base64 = canvas.toDataURL("image/png");
					resolve(base64);
				} catch (error) {
					reject(error);
				}
			};

			img.onerror = function () {
				reject(new Error("Failed to load image"));
			};

			// Try loading as relative path
			img.src = relativePath.startsWith("/")
				? relativePath.substring(1)
				: relativePath;
		} else {
			// Original canvas method for HTTP URLs
			const canvas = document.createElement("canvas");
			const ctx = canvas.getContext("2d");
			const img = new Image();
			img.crossOrigin = "anonymous";

			img.onload = function () {
				canvas.width = img.width;
				canvas.height = img.height;
				ctx.drawImage(img, 0, 0);

				try {
					const base64 = canvas.toDataURL("image/png");
					resolve(base64);
				} catch (error) {
					reject(error);
				}
			};

			img.onerror = function () {
				reject(new Error("Failed to load image"));
			};

			img.src = imgElement.src.startsWith("http")
				? imgElement.src
				: new URL(imgElement.src, window.location.href).href;
		}
	});
}

// Solution 3: Most robust solution - handles both protocols
function convertImageToBase64Robust(imgElement) {
	return new Promise((resolve, reject) => {
		const canvas = document.createElement("canvas");
		const ctx = canvas.getContext("2d");
		const img = new Image();

		// Don't set crossOrigin for file:// URLs
		if (!imgElement.src.startsWith("file://")) {
			img.crossOrigin = "anonymous";
		}

		img.onload = function () {
			canvas.width = img.width;
			canvas.height = img.height;

			try {
				ctx.drawImage(img, 0, 0);
				const base64 = canvas.toDataURL("image/png");
				resolve(base64);
			} catch (error) {
				console.warn("Canvas conversion failed:", error);
				// Fallback: return original src
				resolve(imgElement.src);
			}
		};

		img.onerror = function () {
			console.warn("Image load failed for:", imgElement.src);
			// Fallback: return original src or placeholder
			resolve(imgElement.src);
		};

		// Handle different URL types
		if (imgElement.src.startsWith("file://")) {
			// For file:// URLs, convert to relative path
			const currentDir = window.location.href.substring(
				0,
				window.location.href.lastIndexOf("/")
			);
			const relativePath = imgElement.src
				.replace("file://", "")
				.replace(currentDir.replace("file://", ""), "");
			img.src = relativePath.startsWith("/")
				? relativePath.substring(1)
				: relativePath;
		} else if (imgElement.src.startsWith("http")) {
			img.src = imgElement.src;
		} else {
			// Convert relative path to absolute path
			img.src = new URL(imgElement.src, window.location.href).href;
		}
	});
}

// Updated exportToDocx function with better error handling
async function exportToDocx(element, filename) {
	try {
		// First, process images in the element
		const tempDiv = element.cloneNode(true);
		const images = tempDiv.querySelectorAll("img");

		// Convert all images to base64
		const imagePromises = Array.from(images).map(async (img) => {
			try {
				const base64 = await convertImageToBase64Robust(img);
				img.src = base64;
				return true;
			} catch (error) {
				console.warn("Failed to convert image:", img.src, error);

				// Option 1: Replace with placeholder text
				const placeholder = document.createElement("span");
				placeholder.textContent = "[Image not available]";
				placeholder.style.fontStyle = "italic";
				placeholder.style.color = "#666";
				img.parentNode.replaceChild(placeholder, img);

				return false;
			}
		});

		await Promise.all(imagePromises);

		// Create a complete HTML document with embedded styles
		const htmlContent = `
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <title>${filename}</title>
                <style>
                    body { 
                        font-family: Arial, sans-serif; 
                        line-height: 1.6;
                        color: #333;
                        max-width: 800px;
                        margin: 0 auto;
                        padding: 20px;
                    }
                    h1 { 
                        color: #2c3e50; 
                        border-bottom: 3px solid #3498db;
                        padding-bottom: 10px;
                        font-size: 24px;
                    }
                    h2 { 
                        color: #34495e; 
                        margin-top: 30px;
                        border-bottom: 2px solid #95a5a6;
                        padding-bottom: 5px;
                        font-size: 20px;
                    }
                    h3 {
                        color: #2c3e50;
                        margin-top: 25px;
                        font-size: 16px;
                    }
                    h4 {
                        color: #2c3e50;
                        margin-top: 20px;
                        font-size: 14px;
                    }
                    table { 
                        width: 100%; 
                        border-collapse: collapse; 
                        margin: 20px 0;
                    }
                    th, td { 
                        border: 1px solid #ddd; 
                        padding: 12px; 
                        text-align: left;
                    }
                    th { 
                        background-color: #3498db; 
                        color: white;
                        font-weight: bold;
                    }
                    .highlight {
                        background-color: #fff3cd;
                        padding: 15px;
                        border-radius: 5px;
                        border-left: 4px solid #ffc107;
                        margin: 20px 0;
                    }
                    .info-box {
                        background-color: #d1ecf1;
                        padding: 15px;
                        border-radius: 5px;
                        border-left: 4px solid #17a2b8;
                        margin: 20px 0;
                    }
                    .success-box {
                        background-color: #d4edda;
                        padding: 15px;
                        border-radius: 5px;
                        border-left: 4px solid #28a745;
                        margin: 20px 0;
                    }
                    .feature-grid {
                        display: table;
                        width: 100%;
                    }
                    .feature-item {
                        display: table-cell;
                        vertical-align: top;
                        padding: 15px;
                        border: 1px solid #ddd;
                        margin: 10px 0;
                    }
                    blockquote {
                        border-left: 4px solid #6c757d;
                        padding-left: 20px;
                        margin: 20px 0;
                        font-style: italic;
                        color: #6c757d;
                    }
                    ul, ol { 
                        padding-left: 30px;
                        line-height: 1.6;
                    }
                    li { 
                        margin-bottom: 8px;
                    }
                    pre {
                        background-color: #f8f9fa;
                        padding: 15px;
                        border-radius: 5px;
                        border: 1px solid #e9ecef;
                        overflow-x: auto;
                    }
                    code {
                        background-color: #f8f9fa;
                        padding: 2px 4px;
                        border-radius: 3px;
                        font-family: 'Courier New', monospace;
                    }
                    .mx-tools {
                        display:none
                    }
                    .question-parent {
                        margin: 1rem 0;
                    }
                    /* Image styles for better display in DOCX */
                    img {
                        max-width: 100%;
                        height: auto;
                        display: block;
                        margin: 10px 0;
                    }
                </style>
            </head>
            <body>
                ${tempDiv.innerHTML}
            </body>
            </html>
        `;

		// Convert HTML to DOCX
		const converted = htmlDocx.asBlob(htmlContent, {
			orientation: "portrait",
			margins: {
				top: 720, // 0.5 inch
				right: 720,
				bottom: 720,
				left: 720,
			},
		});

		console.log("Processed HTML Content:", htmlContent);

		// Save the file
		saveAs(converted, `${filename}.docx`);
		// showStatus("success", `Successfully exported ${filename}.docx`);
	} catch (error) {
		console.error("Export error:", error);
		// showStatus("error", "Export failed: " + error.message);
	}
}
