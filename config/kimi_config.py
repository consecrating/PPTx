"""
Kimi AI Configuration for PPTx Generation

Configures Kimi K3 (or K2.7-code) as the AI backend for generating
professional PowerPoint presentations.

Kimi K3 is ideal for PPTx generation because:
- 2.8T parameters with strong coding capabilities
- 1M token context window (can process large templates/content)
- Excellent at long-horizon tasks (multi-slide generation)
- Native tool calling support (tool_choice: auto/none/required)
- Always-on reasoning (produces better slide structure/design decisions)

Setup:
    1. Get API key from https://platform.kimi.ai
    2. Set KIMI_API_KEY environment variable
    3. Use this config with the PPTx generation scripts

Usage:
    from config.kimi_config import get_kimi_client, generate_pptx_content

    client = get_kimi_client()
    content = generate_pptx_content(client, "Create a pitch deck about AI startups")
"""

import os
import json
from typing import Optional


# ============================================================
# KIMI API CONFIGURATION
# ============================================================
#
# FREE ACCESS OPTIONS:
# 1. Kimi Direct (platform.kimi.ai) — Sign up for trial credits
#    - New accounts may receive starter credits
#    - Check your console balance at platform.kimi.ai
#
# 2. OpenRouter (openrouter.ai) — FREE route available!
#    - Model: moonshotai/kimi-k3
#    - Free tier with rate limits, no credit card required
#    - Best free option for API access
#
# 3. Kimi App (kimi.com) — FREE chat tier
#    - Use Kimi K3 directly in browser at kimi.com
#    - Includes Kimi Code (terminal agent) for coding
#    - Rate-limited but genuinely free
#
# 4. Self-host — Open weights available (free, need GPUs)
#    - Weights at: https://huggingface.co/MoonshotAI
#    - Requires multi-GPU setup (2.8T params)
#    - Wait for quantized community builds for smaller setups
# ============================================================

KIMI_CONFIG = {
    # OpenRouter — FREE Kimi K3 access (RECOMMENDED FOR FREE USE)
    "openrouter": {
        "base_url": "https://openrouter.ai/api/v1",
        "env_key": "OPENROUTER_API_KEY",
        "free": True,
        "signup_url": "https://openrouter.ai (free account, no credit card)",
        "models": {
            "moonshotai/kimi-k3": {
                "context_window": 1048576,  # 1M tokens
                "max_output": 131072,
                "supports_vision": True,
                "supports_tools": True,
                "best_for": "FREE Kimi K3 access for PPTx generation",
                "free_tier": True,
                "note": "Free route with rate limits; may throttle under heavy load",
            },
        },
        "default_model": "moonshotai/kimi-k3",
    },

    # Direct Kimi Platform (trial credits on signup)
    "kimi_direct": {
        "base_url": "https://api.moonshot.cn/v1",
        "env_key": "KIMI_API_KEY",
        "free": False,
        "signup_url": "https://platform.kimi.ai (check console for trial credits)",
        "models": {
            "kimi-k3": {
                "context_window": 1048576,  # 1M tokens
                "max_output": 131072,
                "reasoning_effort": "high",  # low/high/max
                "supports_vision": True,
                "supports_tools": True,
                "best_for": "Best quality — complex multi-slide presentations",
            },
            "kimi-k2.7-code": {
                "context_window": 262144,  # 256K tokens
                "max_output": 65536,
                "supports_vision": True,
                "supports_tools": True,
                "best_for": "Fast pptxgenjs code generation",
            },
            "kimi-k2.7-code-highspeed": {
                "context_window": 262144,
                "max_output": 65536,
                "supports_vision": True,
                "supports_tools": True,
                "best_for": "Fastest code generation (same model, higher throughput)",
            },
        },
        "default_model": "kimi-k3",
    },

    # TokenRouter (for K2.x models — K3 not yet available)
    "tokenrouter": {
        "base_url": "https://tokenrouter.me/v1",
        "env_key": "TOKENROUTER_API_KEY",
        "free": False,
        "signup_url": "https://tokenrouter.com",
        "models": {
            "kimi-k2p7-code": {
                "context_window": 262144,
                "max_output": 65536,
                "supports_vision": True,
                "supports_tools": True,
                "best_for": "Agentic coding via TokenRouter",
            },
            "kimi-k2p7-code-fast": {
                "context_window": 262144,
                "max_output": 65536,
                "supports_vision": True,
                "supports_tools": True,
                "best_for": "Fast coding via TokenRouter",
            },
            "kimi-k2p6": {
                "context_window": 262144,
                "max_output": 65536,
                "supports_vision": True,
                "supports_tools": True,
                "best_for": "Long-horizon tasks via TokenRouter",
            },
        },
        "default_model": "kimi-k2p7-code",
    },
}

# Default provider — OpenRouter is FREE
DEFAULT_PROVIDER = "openrouter"


# ============================================================
# CLIENT SETUP
# ============================================================

def get_kimi_client(provider: str = DEFAULT_PROVIDER, api_key: Optional[str] = None):
    """
    Get an OpenAI-compatible client configured for Kimi.

    Args:
        provider: "openrouter" (FREE), "kimi_direct", or "tokenrouter"
        api_key: Override API key (otherwise reads from env)

    Returns:
        OpenAI client instance
    """
    try:
        from openai import OpenAI
    except ImportError:
        raise ImportError("Install openai SDK: pip install openai")

    config = KIMI_CONFIG[provider]
    key = api_key or os.environ.get(config["env_key"])

    if not key:
        raise ValueError(
            f"API key not found. Set {config['env_key']} environment variable "
            f"or pass api_key parameter.\n\n"
            f"Get your FREE key from: {config.get('signup_url', 'the provider website')}\n"
            f"{'(No credit card required!)' if config.get('free') else ''}\n\n"
            f"Setup:\n"
            f"  export {config['env_key']}='your-key-here'\n"
        )

    return OpenAI(base_url=config["base_url"], api_key=key)


def get_model_info(model: Optional[str] = None, provider: str = DEFAULT_PROVIDER) -> dict:
    """Get model configuration info."""
    config = KIMI_CONFIG[provider]
    model = model or config["default_model"]
    return config["models"].get(model, {})


# ============================================================
# PPTX GENERATION SYSTEM PROMPT
# ============================================================

PPTX_SYSTEM_PROMPT = """You are an expert PowerPoint presentation designer and pptxgenjs developer.

Your task is to generate complete, professional Node.js scripts using pptxgenjs that create stunning presentations.

## Critical Rules (pptxgenjs):
- ALWAYS set `pres.layout = "LAYOUT_16x9"` before adding slides (10" x 5.625")
- Hex colors: NEVER use `#` prefix. Use "1E2761" not "#1E2761"
- NEVER use 8-digit hex (no alpha channel in hex)
- Build FRESH option objects for each addText/addShape call (pptxgenjs mutates them)
- Shadow offset must be >= 0 (negative corrupts file). Use angle for direction.
- Use `charSpacing` not `letterSpacing` (letterSpacing is silently ignored)
- Bullet lists: `bullet: true` on each item, `breakLine: true` on all except last
- Space bullets with `paraSpaceAfter`, not `lineSpacing`
- One `new pptxgen()` per output file
- `rectRadius` only works on ROUNDED_RECTANGLE, not RECTANGLE
- Gradient fills NOT supported — use gradient images instead
- Set `margin: 0` on text boxes for precise alignment
- Speaker notes: `slide.addNotes("text")` once per slide
- Stacked bar/column charts: dataLabelPosition must be "ctr", "inEnd", or "inBase" (outEnd corrupts)
- Combo charts with secondaryValAxis MUST define both valAxes AND catAxes (2 entries each)

## Typography (Safe Fonts):
- Sans: Arial, Calibri
- Serif: Cambria, Times New Roman, Bookman Old Style, Century Schoolbook
- Mono: Courier New
- NEVER use Aptos

## Design Principles:
- Every slide needs a visual element (never text-only)
- Vary layouts across slides (don't repeat same layout)
- Left-align body text (center only titles)
- Title: 36-44pt bold, Body: 14-16pt, Caption: 10-12pt
- 0.5" minimum margins, 0.3-0.5" between blocks
- Dark backgrounds for title/conclusion, light for content
- Use ONE color palette consistently throughout
- NO accent lines under titles, NO decorative color bars
- NO beige/cream backgrounds (use white or brand colors)

## Output Format:
Generate a complete, runnable Node.js script that:
1. Requires pptxgenjs
2. Defines a color palette
3. Creates all slides with proper content
4. Includes transitions between slides
5. Saves to a .pptx file
6. Can be run directly with `node script.js`
"""


# ============================================================
# GENERATION FUNCTIONS
# ============================================================

def generate_pptx_content(
    client,
    prompt: str,
    model: Optional[str] = None,
    provider: str = DEFAULT_PROVIDER,
    reasoning_effort: str = "high",
    stream: bool = False,
) -> str:
    """
    Generate pptxgenjs code for a presentation using Kimi.

    Args:
        client: OpenAI client (from get_kimi_client)
        prompt: User's presentation request
        model: Model ID (default: provider's default)
        provider: "kimi_direct" or "tokenrouter"
        reasoning_effort: "low", "high", or "max" (K3 only)
        stream: Whether to stream the response

    Returns:
        Generated pptxgenjs JavaScript code
    """
    config = KIMI_CONFIG[provider]
    model = model or config["default_model"]

    messages = [
        {"role": "system", "content": PPTX_SYSTEM_PROMPT},
        {"role": "user", "content": f"Create a professional PowerPoint presentation:\n\n{prompt}\n\nGenerate the complete pptxgenjs Node.js script."},
    ]

    kwargs = {
        "model": model,
        "messages": messages,
        "stream": stream,
    }

    # K3-specific: reasoning_effort (only for kimi_direct provider)
    if model == "kimi-k3" and provider == "kimi_direct":
        kwargs["reasoning_effort"] = reasoning_effort

    if stream:
        response = client.chat.completions.create(**kwargs)
        content = ""
        for chunk in response:
            if chunk.choices and chunk.choices[0].delta.content:
                content += chunk.choices[0].delta.content
                print(chunk.choices[0].delta.content, end="", flush=True)
        print()
        return content
    else:
        response = client.chat.completions.create(**kwargs)
        return response.choices[0].message.content


def generate_and_save(
    prompt: str,
    output_path: str = "generated_presentation.js",
    model: Optional[str] = None,
    provider: str = DEFAULT_PROVIDER,
    api_key: Optional[str] = None,
    reasoning_effort: str = "high",
) -> str:
    """
    Generate a pptxgenjs script and save to file.

    Args:
        prompt: Description of the presentation to create
        output_path: Where to save the generated .js file
        model: Model to use
        provider: API provider
        api_key: Override API key
        reasoning_effort: K3 reasoning level

    Returns:
        Path to the generated file
    """
    client = get_kimi_client(provider, api_key)

    print(f"Generating presentation with {model or KIMI_CONFIG[provider]['default_model']}...")
    print(f"Provider: {provider}")
    print(f"Prompt: {prompt[:100]}...")
    print()

    code = generate_pptx_content(
        client, prompt, model, provider, reasoning_effort, stream=True
    )

    # Extract JavaScript code from markdown code blocks if present
    if "```javascript" in code:
        code = code.split("```javascript")[1].split("```")[0].strip()
    elif "```js" in code:
        code = code.split("```js")[1].split("```")[0].strip()

    with open(output_path, "w") as f:
        f.write(code)

    print(f"\nSaved to: {output_path}")
    print(f"Run: node {output_path}")
    return output_path


# ============================================================
# CLI INTERFACE
# ============================================================

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Generate PPTX with Kimi AI")
    parser.add_argument("prompt", help="Presentation description")
    parser.add_argument("-o", "--output", default="generated_presentation.js",
                       help="Output file path")
    parser.add_argument("-m", "--model", default=None,
                       help="Model ID (default: kimi-k3)")
    parser.add_argument("-p", "--provider", default=DEFAULT_PROVIDER,
                       choices=["kimi_direct", "tokenrouter"],
                       help="API provider")
    parser.add_argument("-r", "--reasoning", default="high",
                       choices=["low", "high", "max"],
                       help="Reasoning effort (K3 only)")
    parser.add_argument("--list-models", action="store_true",
                       help="List available models")

    args = parser.parse_args()

    if args.list_models:
        print("\nAvailable Models:")
        print("=" * 70)
        for provider_name, provider_config in KIMI_CONFIG.items():
            print(f"\n[{provider_name}] ({provider_config['base_url']})")
            print(f"  API Key env: {provider_config['env_key']}")
            for model_id, model_info in provider_config["models"].items():
                default = " (DEFAULT)" if model_id == provider_config["default_model"] else ""
                print(f"  - {model_id}{default}")
                print(f"    Context: {model_info['context_window']:,} tokens")
                print(f"    Best for: {model_info['best_for']}")
    else:
        generate_and_save(
            prompt=args.prompt,
            output_path=args.output,
            model=args.model,
            provider=args.provider,
            reasoning_effort=args.reasoning,
        )
