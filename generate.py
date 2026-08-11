#!/usr/bin/env python3
"""
PPTx AI Generator — Create professional presentations with Kimi K3

Quick usage:
    python generate.py "A 10-slide pitch deck for an AI startup"
    python generate.py "Quarterly business review for Q2 2026" -m kimi-k3 -r max
    python generate.py "Product launch presentation" --provider tokenrouter

Requirements:
    pip install openai
    export KIMI_API_KEY="sk-your-key"  (from https://platform.kimi.ai)
"""

import sys
import os

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config.kimi_config import generate_and_save, KIMI_CONFIG

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Generate professional PPTX presentations with Kimi AI",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python generate.py "A startup pitch deck about AI-powered healthcare"
  python generate.py "Quarterly revenue report Q2 2026" -o quarterly.js
  python generate.py "Team onboarding presentation" -p kimi_direct -r max
  python generate.py "Product roadmap 2026" --provider openrouter
  python generate.py --list-models

FREE Setup (OpenRouter — no credit card needed):
  1. Sign up at https://openrouter.ai (free)
  2. export OPENROUTER_API_KEY='sk-or-...'
  3. python generate.py "Your presentation idea"
        """
    )
    parser.add_argument("prompt", nargs="?", help="What presentation to create")
    parser.add_argument("-o", "--output", default="generated_presentation.js",
                       help="Output JS file path (default: generated_presentation.js)")
    parser.add_argument("-m", "--model", default=None,
                       help="Model ID (default: kimi-k3)")
    parser.add_argument("-p", "--provider", default=os.environ.get("PPTX_PROVIDER", "openrouter"),
                       choices=["openrouter", "kimi_direct", "tokenrouter"],
                       help="API provider (default: openrouter — FREE)")
    parser.add_argument("-r", "--reasoning", default=os.environ.get("PPTX_REASONING_EFFORT", "high"),
                       choices=["low", "high", "max"],
                       help="Reasoning effort — K3 only (default: high)")
    parser.add_argument("--list-models", action="store_true",
                       help="List all available models and exit")

    args = parser.parse_args()

    if args.list_models:
        print("\n╔══════════════════════════════════════════════════════════════╗")
        print("║          PPTx Generator — Available AI Models               ║")
        print("╚══════════════════════════════════════════════════════════════╝")
        for provider_name, config in KIMI_CONFIG.items():
            env_key = config["env_key"]
            has_key = "✓" if os.environ.get(env_key) else "✗"
            print(f"\n┌─ [{provider_name}] {config['base_url']}")
            print(f"│  Key ({env_key}): {has_key}")
            for model_id, info in config["models"].items():
                default = " ★ DEFAULT" if model_id == config["default_model"] else ""
                ctx = f"{info['context_window']:,}"
                print(f"│  • {model_id:<30} {ctx:>12} ctx  │ {info['best_for']}{default}")
            print("└" + "─" * 64)
        print("\nSet your API key:")
        print("  export KIMI_API_KEY='sk-...'       # from https://platform.kimi.ai")
        print("  export TOKENROUTER_API_KEY='sk-...' # from https://tokenrouter.com")
        sys.exit(0)

    if not args.prompt:
        parser.print_help()
        print("\nError: Please provide a presentation prompt.")
        sys.exit(1)

    try:
        output = generate_and_save(
            prompt=args.prompt,
            output_path=args.output,
            model=args.model,
            provider=args.provider,
            reasoning_effort=args.reasoning,
        )
        print(f"\n{'='*60}")
        print(f"  Next steps:")
        print(f"    1. node {output}                    # Create .pptx")
        print(f"    2. python scripts/office/validate.py *.pptx  # Validate")
        print(f"{'='*60}")
    except ValueError as e:
        print(f"\nError: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\nError: {e}")
        sys.exit(1)
