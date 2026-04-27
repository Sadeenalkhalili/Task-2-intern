import os
import re
import shutil
import zipfile
import tempfile
import argparse
from pathlib import Path
from html import escape

import requests
from lxml import etree


ARABIC_RE = re.compile(r"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF]")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": W_NS}


def has_arabic(text: str) -> bool:
    return bool(text) and bool(ARABIC_RE.search(text))


def get_api_key() -> str:
    key = os.environ.get("DEEPL_API_KEY")
    if not key:
        raise RuntimeError("DEEPL_API_KEY is not set.")
    return key


def get_base_url(use_free_api: bool) -> str:
    return "https://api-free.deepl.com" if use_free_api else "https://api.deepl.com"


def unzip_docx(input_path: str, work_dir: str) -> None:
    with zipfile.ZipFile(input_path, "r") as zip_ref:
        zip_ref.extractall(work_dir)


def zip_docx(work_dir: str, output_path: str) -> None:
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as docx_zip:
        for file_path in Path(work_dir).rglob("*"):
            if file_path.is_file():
                relative_path = file_path.relative_to(work_dir)
                docx_zip.write(file_path, relative_path.as_posix())


def get_word_xml_files(work_dir: str) -> list[Path]:
    word_dir = Path(work_dir) / "word"

    wanted_patterns = [
        "document.xml",
        "header*.xml",
        "footer*.xml",
        "footnotes.xml",
        "endnotes.xml",
        "comments.xml",
    ]

    files = []
    for pattern in wanted_patterns:
        files.extend(word_dir.glob(pattern))

    return files


def paragraph_to_deepl_xml(paragraph, paragraph_id: int) -> tuple[str, list]:
    """
    Convert one Word paragraph into small XML for DeepL.

    Each <w:t> text node becomes:
    <r id="0">Arabic text</r>

    DeepL translates the text but keeps the <r id=""> tags.
    Then we put each translated result back into the original <w:t>.
    """
    text_nodes = paragraph.xpath(".//w:t", namespaces=NSMAP)

    parts = [f'<p id="{paragraph_id}">']
    used_nodes = []

    run_id = 0
    for node in text_nodes:
        original_text = node.text or ""

        if original_text == "":
            continue

        used_nodes.append(node)
        parts.append(f'<r id="{run_id}">{escape(original_text)}</r>')
        run_id += 1

    parts.append("</p>")

    return "".join(parts), used_nodes


def parse_translated_deepl_xml(translated_xml: str) -> dict[int, str]:
    """
    Read DeepL's returned XML and extract translated text by run id.
    """
    root = etree.fromstring(translated_xml.encode("utf-8"))
    result = {}

    for r in root.xpath(".//r"):
        run_id = int(r.get("id"))
        translated_text = "".join(r.itertext())
        result[run_id] = translated_text

    return result


class DeepLTextTranslator:
    def __init__(
        self,
        api_key: str,
        base_url: str,
        target_lang: str = "EN",
        source_lang: str | None = "AR",
    ):
        self.api_key = api_key
        self.base_url = base_url
        self.target_lang = target_lang
        self.source_lang = source_lang
        self.cache = {}

    def translate_xml_batch(self, xml_items: list[str]) -> list[str]:
        """
        Translate a batch of XML fragments using DeepL text API.
        """
        if not xml_items:
            return []

        translated_results = []

        for item in xml_items:
            if item in self.cache:
                translated_results.append(self.cache[item])
                continue

            payload = {
                "text": [item],
                "target_lang": self.target_lang,
                "tag_handling": "xml",
                "tag_handling_version": "v2",
                "preserve_formatting": True,
            }

            if self.source_lang:
                payload["source_lang"] = self.source_lang

            headers = {
                "Authorization": f"DeepL-Auth-Key {self.api_key}",
                "Content-Type": "application/json",
            }

            response = requests.post(
                f"{self.base_url}/v2/translate",
                headers=headers,
                json=payload,
                timeout=60,
            )

            response.raise_for_status()
            data = response.json()

            translated = data["translations"][0]["text"]
            self.cache[item] = translated
            translated_results.append(translated)

        return translated_results


def translate_xml_file(xml_path: Path, translator: DeepLTextTranslator) -> int:
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    tree = etree.parse(str(xml_path), parser)
    root = tree.getroot()
    remove_manual_page_breaks(root)

    paragraphs = root.xpath(".//w:p", namespaces=NSMAP)

    items_to_translate = []
    nodes_for_items = []

    paragraph_counter = 0

    for paragraph in paragraphs:
        full_text = "".join(paragraph.xpath(".//w:t/text()", namespaces=NSMAP))

        if not has_arabic(full_text):
            continue

        deepl_xml, text_nodes = paragraph_to_deepl_xml(paragraph, paragraph_counter)

        if not text_nodes:
            continue

        items_to_translate.append(deepl_xml)
        nodes_for_items.append(text_nodes)
        paragraph_counter += 1

    translated_items = translator.translate_xml_batch(items_to_translate)

    changed_count = 0

    for translated_xml, original_nodes in zip(translated_items, nodes_for_items):
        try:
            translated_by_id = parse_translated_deepl_xml(translated_xml)

            for idx, node in enumerate(original_nodes):
                node.text = clean_translation_text(translated_by_id.get(idx, ""))

            force_ltr_paragraph(node.getparent().getparent())

            changed_count += 1

        except Exception as error:
            print(f"Warning: Could not parse translated XML in {xml_path.name}: {error}")

    tree.write(
        str(xml_path),
        xml_declaration=True,
        encoding="UTF-8",
        standalone=None,
    )

    return changed_count

def clean_translation_text(text: str) -> str:
    if not text:
        return text

    text = text.replace("“", '"').replace("”", '"')
    text = text.replace("‘", "'").replace("’", "'")

    # Remove quotes around short translated terms
    text = re.sub(r'"([^"]{1,40})"', r"\1", text)

    return text

def remove_manual_page_breaks(root):
    for br in root.xpath(".//w:br[@w:type='page']", namespaces=NSMAP):
        parent = br.getparent()
        if parent is not None:
            parent.remove(br)
            
def translate_docx_xml(
    input_path: str,
    output_path: str,
    target_lang: str = "EN",
    source_lang: str | None = "AR",
    use_free_api: bool = True,
) -> None:
    if not input_path.lower().endswith(".docx"):
        raise ValueError("Input file must be .docx")

    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    api_key = get_api_key()
    base_url = get_base_url(use_free_api)

    translator = DeepLTextTranslator(
        api_key=api_key,
        base_url=base_url,
        target_lang=target_lang,
        source_lang=source_lang,
    )

    with tempfile.TemporaryDirectory() as work_dir:
        print("Unzipping DOCX...")
        unzip_docx(input_path, work_dir)

        xml_files = get_word_xml_files(work_dir)
        print(f"Found {len(xml_files)} Word XML files.")

        total_changed = 0

        for xml_file in xml_files:
            print(f"Processing: {xml_file.name}")
            changed = translate_xml_file(xml_file, translator)
            total_changed += changed

        print("Rebuilding DOCX...")
        zip_docx(work_dir, output_path)

    print(f"Done. Translated {total_changed} paragraphs.")
    print(f"Saved to: {output_path}")


def parse_args():
    parser = argparse.ArgumentParser(
        description="Translate Arabic DOCX to English while preserving DOCX XML structure."
    )

    parser.add_argument("--input", required=True, help="Input .docx file")
    parser.add_argument("--output", required=True, help="Output .docx file")
    parser.add_argument("--target-lang", default="EN", help="Target language, default EN")
    parser.add_argument("--source-lang", default="AR", help="Source language, default AR")
    parser.add_argument("--pro", action="store_true", help="Use DeepL Pro instead of Free")

    return parser.parse_args()

def force_ltr_paragraph(paragraph):
    pPr = paragraph.find("w:pPr", namespaces=NSMAP)
    if pPr is None:
        pPr = etree.Element(f"{{{W_NS}}}pPr")
        paragraph.insert(0, pPr)

    # Remove paragraph RTL/bidi
    for bidi in pPr.xpath("./w:bidi", namespaces=NSMAP):
        pPr.remove(bidi)

    jc = pPr.find("w:jc", namespaces=NSMAP)
    if jc is None:
        jc = etree.SubElement(pPr, f"{{{W_NS}}}jc")
    jc.set(f"{{{W_NS}}}val", "left")

    # Force all runs to LTR
    for run in paragraph.xpath(".//w:r", namespaces=NSMAP):
        rPr = run.find("w:rPr", namespaces=NSMAP)
        if rPr is None:
            rPr = etree.Element(f"{{{W_NS}}}rPr")
            run.insert(0, rPr)

        for rtl in rPr.xpath("./w:rtl", namespaces=NSMAP):
            rPr.remove(rtl)
            
def main():
    args = parse_args()

    translate_docx_xml(
        input_path=args.input,
        output_path=args.output,
        target_lang=args.target_lang.upper(),
        source_lang=args.source_lang.upper() if args.source_lang else None,
        use_free_api=not args.pro,
    )


if __name__ == "__main__":
    main()