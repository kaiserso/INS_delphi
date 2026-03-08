#!/usr/bin/env python3
"""
Generate QR code for the Delphi gateway URL.

Usage:
    python code/generate_qrcode.py [--output OUTPUT_FILE] [--url URL]

Examples:
    # Generate QR code with default settings
    python code/generate_qrcode.py
    
    # Specify custom output file
    python code/generate_qrcode.py --output gateway_qr.png
    
    # Use custom URL
    python code/generate_qrcode.py --url https://example.com
"""

import argparse
import qrcode
from pathlib import Path


DEFAULT_GATEWAY_URL = "https://kaiserso.github.io/INS_delphi/malaria/gateway.html"
DEFAULT_OUTPUT = "gateway_qr.png"


def generate_qr_code(url: str, output_path: str, box_size: int = 10, border: int = 4):
    """
    Generate a QR code for the given URL.
    
    Args:
        url: The URL to encode in the QR code
        output_path: Path where the QR code image will be saved
        box_size: Size of each box in pixels (default: 10)
        border: Border size in boxes (default: 4, minimum is 4)
    """
    # Create QR code instance
    qr = qrcode.QRCode(
        version=1,  # Auto-adjust size
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=box_size,
        border=border,
    )
    
    # Add data
    qr.add_data(url)
    qr.make(fit=True)
    
    # Create image
    img = qr.make_image(fill_color="black", back_color="white")
    
    # Save image
    output_file = Path(output_path)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    img.save(output_file)
    
    print(f"✓ QR code generated successfully!")
    print(f"  URL: {url}")
    print(f"  Saved to: {output_file.resolve()}")
    
    return output_file


def main():
    parser = argparse.ArgumentParser(
        description="Generate QR code for Delphi gateway URL",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python code/generate_qrcode.py
  python code/generate_qrcode.py --output qr_codes/malaria_gateway.png
  python code/generate_qrcode.py --url https://example.com --box-size 15
        """
    )
    
    parser.add_argument(
        "--url",
        default=DEFAULT_GATEWAY_URL,
        help=f"URL to encode in QR code (default: {DEFAULT_GATEWAY_URL})"
    )
    
    parser.add_argument(
        "--output",
        default=DEFAULT_OUTPUT,
        help=f"Output file path (default: {DEFAULT_OUTPUT})"
    )
    
    parser.add_argument(
        "--box-size",
        type=int,
        default=10,
        help="Size of each box in pixels (default: 10)"
    )
    
    parser.add_argument(
        "--border",
        type=int,
        default=4,
        help="Border size in boxes (default: 4, minimum is 4)"
    )
    
    args = parser.parse_args()
    
    generate_qr_code(
        url=args.url,
        output_path=args.output,
        box_size=args.box_size,
        border=args.border
    )


if __name__ == "__main__":
    main()
