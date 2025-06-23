import requests
import urllib.parse
import json
import os


def swissgerman_to_english(prompt, mapping_file="swissgerman_mapping.json"):
    """
    Load Swiss German to English mappings from file and apply them word by word.
    
    Args:
        prompt (str): The Swiss German text to translate
        mapping_file (str): Path to the JSON mapping file
        
    Returns:
        str: Translated text with Swiss German words replaced by English equivalents
        
    Raises:
        FileNotFoundError: If the mapping file doesn't exist
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
    Generates an AI image using Pollinations API.
    Swiss German prompts are translated for AI generation using a local mapping file.
    
    Args:
        prompt (str): The input image prompt (can be Swiss German)
        output_file (str): Image filename to save
        size (str): Image dimensions in format "widthxheight"
        translate_if_needed (bool): If True, Swiss German prompts are mapped to English
        mapping_file (str): Path to Swiss German mapping JSON
        
    Returns:
        str: Path to saved image file, or None if generation failed
    """
    try:
        original_prompt = prompt.strip()
        prompt = swissgerman_to_english(original_prompt, mapping_file)
        prompt = "generiere ein gericht: " + prompt
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