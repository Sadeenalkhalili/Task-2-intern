import os
import time
import argparse
import requests


def get_api_key() -> str:
    api_key = os.environ.get("DEEPL_API_KEY")
    if not api_key:
        raise RuntimeError(
            "DEEPL_API_KEY is not set. "
            "Set it in your environment before running the script."
        )
    return api_key


def get_base_url(use_free_api: bool) -> str:
    if use_free_api:
        return "https://api-free.deepl.com"
    return "https://api.deepl.com"


def upload_document(
    input_path: str,
    target_lang: str,
    api_key: str,
    base_url: str,
    source_lang: str | None = None,
) -> tuple[str, str]:
    upload_url = f"{base_url}/v2/document"

    with open(input_path, "rb") as file_obj:
        files = {
            "file": (
                os.path.basename(input_path),
                file_obj,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        }

        data = {
            "target_lang": target_lang,
        }

        if source_lang:
            data["source_lang"] = source_lang

        headers = {
            "Authorization": f"DeepL-Auth-Key {api_key}",
        }

        response = requests.post(
            upload_url,
            headers=headers,
            data=data,
            files=files,
            timeout=120,
        )
        response.raise_for_status()

        result = response.json()
        return result["document_id"], result["document_key"]


def check_status(
    document_id: str,
    document_key: str,
    api_key: str,
    base_url: str,
) -> dict:
    status_url = f"{base_url}/v2/document/{document_id}"

    headers = {
        "Authorization": f"DeepL-Auth-Key {api_key}",
    }

    data = {
        "document_key": document_key,
    }

    response = requests.post(
        status_url,
        headers=headers,
        data=data,
        timeout=60,
    )
    response.raise_for_status()
    return response.json()


def wait_until_done(
    document_id: str,
    document_key: str,
    api_key: str,
    base_url: str,
    poll_interval: int = 5,
) -> None:
    while True:
        result = check_status(document_id, document_key, api_key, base_url)
        status = result.get("status", "")

        print(f"Status: {status}")

        if status == "done":
            return

        if status == "error":
            message = result.get("message", "Unknown error")
            raise RuntimeError(f"Translation failed: {message}")

        seconds_remaining = result.get("seconds_remaining")
        if isinstance(seconds_remaining, int) and seconds_remaining > 0:
            time.sleep(min(seconds_remaining, poll_interval))
        else:
            time.sleep(poll_interval)


def download_document(
    document_id: str,
    document_key: str,
    output_path: str,
    api_key: str,
    base_url: str,
) -> None:
    download_url = f"{base_url}/v2/document/{document_id}/result"

    headers = {
        "Authorization": f"DeepL-Auth-Key {api_key}",
    }

    data = {
        "document_key": document_key,
    }

    response = requests.post(
        download_url,
        headers=headers,
        data=data,
        timeout=120,
    )
    response.raise_for_status()

    with open(output_path, "wb") as file_obj:
        file_obj.write(response.content)


def validate_input_file(input_path: str) -> None:
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if not input_path.lower().endswith(".docx"):
        raise ValueError("Input file must be a .docx file")


def translate_docx(
    input_path: str,
    output_path: str,
    target_lang: str,
    use_free_api: bool,
    source_lang: str | None = None,
) -> None:
    validate_input_file(input_path)

    api_key = get_api_key()
    base_url = get_base_url(use_free_api)

    print("Uploading document...")
    document_id, document_key = upload_document(
        input_path=input_path,
        target_lang=target_lang,
        api_key=api_key,
        base_url=base_url,
        source_lang=source_lang,
    )

    print("Waiting for translation to finish...")
    wait_until_done(
        document_id=document_id,
        document_key=document_key,
        api_key=api_key,
        base_url=base_url,
    )

    print("Downloading translated document...")
    download_document(
        document_id=document_id,
        document_key=document_key,
        output_path=output_path,
        api_key=api_key,
        base_url=base_url,
    )

    print(f"Done. Saved translated file to: {output_path}")


def build_output_name(input_path: str, target_lang: str) -> str:
    base, ext = os.path.splitext(input_path)
    return f"{base}_{target_lang.lower()}{ext}"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Translate a DOCX file using the DeepL document translation API."
    )

    parser.add_argument(
        "--input",
        required=True,
        help="Path to the input .docx file",
    )

    parser.add_argument(
        "--output",
        required=False,
        help="Path to save the translated .docx file",
    )

    parser.add_argument(
        "--target-lang",
        default="EN",
        help="Target language code, default is EN",
    )

    parser.add_argument(
        "--source-lang",
        default=None,
        help="Optional source language code, for example AR",
    )

    parser.add_argument(
        "--pro",
        action="store_true",
        help="Use DeepL Pro API endpoint instead of Free API endpoint",
    )

    return parser.parse_args()


def main() -> None:
    args = parse_args()

    input_path = args.input
    output_path = args.output or build_output_name(input_path, args.target_lang)
    target_lang = args.target_lang.upper()
    source_lang = args.source_lang.upper() if args.source_lang else None
    use_free_api = not args.pro

    try:
        translate_docx(
            input_path=input_path,
            output_path=output_path,
            target_lang=target_lang,
            use_free_api=use_free_api,
            source_lang=source_lang,
        )
    except Exception as error:
        print(f"Error: {error}")


if __name__ == "__main__":
    main()