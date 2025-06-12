import requests
import urllib.parse
import json
import os


def swissgerman_to_english(prompt, mapping_file="swissgerman_mapping.json"):
    """
    Load Swiss German to English mappings from file and apply them word by word.
    """
    if not os.path.exists(mapping_file):
        raise FileNotFoundError(f"Mapping file '{mapping_file}' not found.")

    with open(mapping_file, 'r', encoding='utf-8') as f:
        replacements = json.load(f)

    words = prompt.lower().split()
    translated = [replacements.get(word, word) for word in words]
    return " ".join(translated)


def generate_image_best(
    prompt,
    output_file="image.png",
    size="1024x1024",
    use_pexels=False,
    pexels_api_key=None,
    use_google=False,
    google_api_key=None,
    google_cx=None,
    translate_if_needed=True,
    mapping_file="dict.json"
):
    """
    Generates an AI image, or fetches the first result from Pexels or Google.
    Swiss German prompts are translated for AI generation using a local mapping file.

    Parameters:
    - prompt: The input image prompt (can be Swiss German).
    - output_file: Image filename to save.
    - size: Only used for AI images (widthxheight).
    - use_pexels: If True, fetch image from Pexels API.
    - pexels_api_key: Required if using Pexels.
    - use_google: If True, fetch image from Google Custom Search.
    - google_api_key / google_cx: Required if using Google.
    - translate_if_needed: If True, Swiss German prompts are mapped to English.
    - mapping_file: Path to Swiss German mapping JSON.
    """
    try:
        original_prompt = prompt.strip()

        if not use_google and not use_pexels and translate_if_needed:
            prompt = swissgerman_to_english(original_prompt, mapping_file)

        encoded_prompt = urllib.parse.quote(prompt)

        if use_pexels:
            if not pexels_api_key:
                raise ValueError("Pexels API key is required for Pexels search.")
            headers = {"Authorization": pexels_api_key}
            search_url = f"https://api.pexels.com/v1/search?query={encoded_prompt}&per_page=1"
            res = requests.get(search_url, headers=headers)
            res.raise_for_status()
            photos = res.json().get("photos")
            if not photos:
                print("No results found on Pexels.")
                return None
            img_url = photos[0]["src"]["original"]
            img_data = requests.get(img_url).content

        elif use_google:
            if not (google_api_key and google_cx):
                raise ValueError("Google API key and CX are required for Google search.")
            search_url = (
                f"https://www.googleapis.com/customsearch/v1?"
                f"q={encoded_prompt}&searchType=image&num=1&key={google_api_key}&cx={google_cx}"
            )
            res = requests.get(search_url)
            res.raise_for_status()
            items = res.json().get("items")
            if not items:
                print("No results found on Google.")
                return None
            img_url = items[0]["link"]
            img_data = requests.get(img_url).content

        else:
            width, height = size.split("x")
            url = f"https://image.pollinations.ai/prompt/{encoded_prompt}?width={width}&height={height}&nologo=true"
            res = requests.get(url, timeout=30)
            if res.status_code != 200:
                print("AI image generation failed.")
                return None
            img_data = res.content

        with open(output_file, "wb") as f:
            f.write(img_data)

        return output_file

    except Exception as e:
        print(f"[Error] {e}")
        return None
