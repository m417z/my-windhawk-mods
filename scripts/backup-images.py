import re
import requests
from pathlib import Path
from typing import List

url_pattern = r"https://i\.imgur\.com/\w+\.\w+"
script_dir = Path(__file__).parent
code_folder_path = script_dir.parent / "mods"
images_folder_path = script_dir.parent / "images"

session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0'})


def find_image_urls(file_path: Path) -> List[str]:
    with file_path.open('r', encoding='utf-8') as file:
        content = file.read()
    return re.findall(url_pattern, content)


def download_image(url: str, save_path: Path):
    if save_path.exists():
        return

    response = session.get(url)
    response.raise_for_status()

    with save_path.open('wb') as file:
        file.write(response.content)

    print(f"Downloaded and saved: {save_path.name}")


def process_code_files(code_folder: Path, images_folder: Path):
    stale_images = list(images_folder.glob('*'))

    image_urls: set[str] = set()
    for file_path in code_folder.glob('*.wh.cpp'):
        image_urls.update(find_image_urls(file_path))

    for url in image_urls:
        # Extract image name from URL
        image_name = url.split('/')[-1]
        image_path = images_folder / image_name
        download_image(url, image_path)

        if image_path in stale_images:
            stale_images.remove(image_path)

    if stale_images:
        print("Stale images:")
        print("\n".join(p.name for p in stale_images))


if __name__ == "__main__":
    process_code_files(code_folder_path, images_folder_path)
