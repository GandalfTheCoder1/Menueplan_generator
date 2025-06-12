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
    translate_if_needed=True,
    mapping_file="dict.json"
):
    """
    Generates an AI image, or fetches the first result from Pexels.
    Swiss German prompts are translated for AI generation using a local mapping file.
    
    Parameters:
    - prompt: The input image prompt (can be Swiss German).
    - output_file: Image filename to save.
    - size: Only used for AI images (widthxheight).
    - use_pexels: If True, fetch image from Pexels API.
    - pexels_api_key: Required if using Pexels.
    - translate_if_needed: If True, Swiss German prompts are mapped to English.
    - mapping_file: Path to Swiss German mapping JSON.
    """
    try:
        original_prompt = prompt.strip()
        
        
        prompt = swissgerman_to_english(original_prompt, mapping_file)
        
        encoded_prompt = urllib.parse.quote(prompt)
        
        
        # AI image generation using Pollinations
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