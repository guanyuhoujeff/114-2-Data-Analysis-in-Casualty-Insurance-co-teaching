#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "google-genai>=1.0.0",
#     "pillow>=10.0.0",
# ]
# ///
"""Generate image using Gemini 2.0 Flash (image generation capable)."""

import sys
import os
from pathlib import Path
from io import BytesIO

from google import genai
from google.genai import types
from PIL import Image as PILImage

API_KEY = os.environ.get("GEMINI_API_KEY", "")
if not API_KEY:
    print("Error: Set GEMINI_API_KEY", file=sys.stderr)
    sys.exit(1)

client = genai.Client(api_key=API_KEY)

prompt = sys.argv[1]
filename = sys.argv[2]

print(f"Generating: {filename} ...")

response = client.models.generate_content(
    model="gemini-2.5-flash-image",
    contents=prompt,
    config=types.GenerateContentConfig(
        response_modalities=["TEXT", "IMAGE"],
    )
)

image_saved = False
for part in response.parts:
    if part.text is not None:
        print(f"  Text: {part.text}")
    elif part.inline_data is not None:
        image_data = part.inline_data.data
        if isinstance(image_data, str):
            import base64
            image_data = base64.b64decode(image_data)
        image = PILImage.open(BytesIO(image_data))
        if image.mode != 'RGB':
            image = image.convert('RGB')
        output_path = Path(filename)
        image.save(str(output_path), 'PNG')
        print(f"  Saved: {output_path.resolve()} ({image.size[0]}x{image.size[1]})")
        image_saved = True

if not image_saved:
    print("Error: No image generated", file=sys.stderr)
    sys.exit(1)
