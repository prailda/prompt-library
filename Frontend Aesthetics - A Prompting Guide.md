# Frontend Aesthetics: A Prompting Guide

Claude can generate high-quality frontends, but without guidance it tends toward generic, conservative designs. This guide shows you how to prompt Claude to produce more distinctive, polished output.

## Prompting for Better Outputs

Claude has strong knowledge of design principles, typography, and color theory, but defaults to safe choices unless explicitly encouraged otherwise. Through experimentation, we've found three strategies that consistently produce better results:

1. **Guide specific design dimensions** - Direct Claude's attention to typography, color, motion, and backgrounds individually
2. **Reference design inspirations** - Suggest sources like IDE themes or cultural aesthetics without being overly prescriptive
3. **Call out common defaults** - Explicitly tell Claude to avoid its tendency toward generic choices

The prompt below applies these strategies across four key design areas.

## The Prompt

To implement these changes, you can append this prompt section to your system prompt or CLAUDE.md file.

DISTILLED_AESTHETICS_PROMPT = """
<frontend_aesthetics>
You tend to converge toward generic, "on distribution" outputs. In frontend design, this creates what users call the "AI slop" aesthetic. Avoid this: make creative, distinctive frontends that surprise and delight. Focus on:

Typography: Choose fonts that are beautiful, unique, and interesting. Avoid generic fonts like Arial and Inter; opt instead for distinctive choices that elevate the frontend's aesthetics.

Color & Theme: Commit to a cohesive aesthetic. Use CSS variables for consistency. Dominant colors with sharp accents outperform timid, evenly-distributed palettes. Draw from IDE themes and cultural aesthetics for inspiration.

Motion: Use animations for effects and micro-interactions. Prioritize CSS-only solutions for HTML. Use Motion library for React when available. Focus on high-impact moments: one well-orchestrated page load with staggered reveals (animation-delay) creates more delight than scattered micro-interactions. 

Backgrounds: Create atmosphere and depth rather than defaulting to solid colors. Layer CSS gradients, use geometric patterns, or add contextual effects that match the overall aesthetic.

Avoid generic AI-generated aesthetics:
- Overused font families (Inter, Roboto, Arial, system fonts)
- Clichéd color schemes (particularly purple gradients on white backgrounds)
- Predictable layouts and component patterns
- Cookie-cutter design that lacks context-specific character

Interpret creatively and make unexpected choices that feel genuinely designed for the context. Vary between light and dark themes, different fonts, different aesthetics. You still tend to converge on common choices (Space Grotesk, for example) across generations. Avoid this: it is critical that you think outside the box!
</frontend_aesthetics>
"""

## Results

Here are the results of UI generations both with and without the prompt section above.

Without guidance, Claude often defaults to simplistic designs with white and purple backgrounds. With the aesthetics prompt, it produces more varied and visually interesting designs.

### Example 1: SaaS Landing Page

**Prompt:** `"Create a SaaS landing page for a project management tool"`

|   |   |
|---|---|
|**Without Aesthetics Prompt**<br><br>![Baseline output without aesthetics guidance](https://raw.githubusercontent.com/anthropics/claude-cookbooks/001e5ca1e735563cdaf9ee5c06019a6f608fd403/images/frontend_aesthetics/baseline_saas.png)|**With Aesthetics Prompt**<br><br>![Enhanced output with distilled aesthetics prompt](https://raw.githubusercontent.com/anthropics/claude-cookbooks/001e5ca1e735563cdaf9ee5c06019a6f608fd403/images/frontend_aesthetics/distilled_saas.png)|

### Example 2: Blog Post

**Prompt:** `"Build a blog post layout with author bio, reading time, and related articles"`

|   |   |
|---|---|
|**Without Aesthetics Prompt**<br><br>![Baseline portfolio without aesthetics guidance](https://raw.githubusercontent.com/anthropics/claude-cookbooks/001e5ca1e735563cdaf9ee5c06019a6f608fd403/images/frontend_aesthetics/baseline_portfolio.png)|**With Aesthetics Prompt**<br><br>![Enhanced portfolio with distilled aesthetics prompt](https://raw.githubusercontent.com/anthropics/claude-cookbooks/001e5ca1e735563cdaf9ee5c06019a6f608fd403/images/frontend_aesthetics/distilled_portfolio.png)|

### Example 3: Admin Table

**Prompt:** `"Create an admin panel with a data table showing users, their roles, and action buttons"`

|   |   |
|---|---|
|**Without Aesthetics Prompt**<br><br>![Baseline dashboard without aesthetics guidance](https://raw.githubusercontent.com/anthropics/claude-cookbooks/001e5ca1e735563cdaf9ee5c06019a6f608fd403/images/frontend_aesthetics/baseline_dashboard.png)|**With Aesthetics Prompt**<br><br>![Enhanced dashboard with distilled aesthetics prompt](https://raw.githubusercontent.com/anthropics/claude-cookbooks/001e5ca1e735563cdaf9ee5c06019a6f608fd403/images/frontend_aesthetics/distilled_dashboard.png)|

## Try It Yourself

First, set up the helper functions:

import os
import re
import webbrowser
import time
import html
from pathlib import Path
from datetime import datetime
from anthropic import Anthropic
from IPython.display import display, HTML as DisplayHTML

client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

def save_html(html_content):
    os.makedirs("html_outputs", exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filepath = f"html_outputs/{timestamp}.html"
    with open(filepath, "w") as f:
        f.write(html_content)
    return filepath

def extract_html(text):
    pattern = r"```(?:html)?\s*(.*?)\s*```"
    matches = re.findall(pattern, text, re.DOTALL)
    return matches[0] if matches else None

def open_in_browser(filepath):
    abs_path = Path(filepath).resolve()
    webbrowser.open(f"file://{abs_path}")
    print(f"🌐 Opened in browser: {filepath}")

def generate_html_with_claude(system_prompt, user_prompt):
    print("🚀 Generating HTML...\n")

    full_response = ""
    start_time = time.time()
    display_id = display(DisplayHTML(""), display_id=True)

    with client.messages.stream(
        model="claude-sonnet-4-5-20250929",
        max_tokens=64000,
        system=system_prompt,
        messages=[{"role": "user", "content": user_prompt}],
    ) as stream:
        for text in stream.text_stream:
            full_response += text
            escaped_text = html.escape(full_response)
            display_html = f"""
            <div id="stream-container" style="border: 2px solid #667eea; border-radius: 8px; padding: 16px; background: #f8f9fa; max-height: 500px; overflow-y: auto;">
                <pre style="margin: 0; font-family: monospace; font-size: 12px; color: #2d2d2d; white-space: pre-wrap; word-wrap: break-word;">{escaped_text}</pre>
            </div>
            <script>
                requestAnimationFrame(() => {{
                    const container = document.getElementById('stream-container');
                    if (container) {{
                        container.scrollTop = container.scrollHeight;
                    }}
                }});
            </script>
            """
            display_id.update(DisplayHTML(display_html))

    elapsed = time.time() - start_time
    escaped_text = html.escape(full_response)
    final_html = f"""
    <div style="border: 2px solid #28a745; border-radius: 8px; padding: 16px; background: #f8f9fa; max-height: 500px; overflow-y: auto;">
        <pre style="margin: 0; font-family: monospace; font-size: 12px; color: #2d2d2d; white-space: pre-wrap; word-wrap: break-word;">{escaped_text}</pre>
    </div>
    """
    display_id.update(DisplayHTML(final_html))

    print(f"\n✅ Complete in {elapsed:.1f}s\n")

    html_content = extract_html(full_response)
    if html_content is None:
        print("❌ Error: Could not extract HTML from response.")
        raise ValueError("Failed to extract HTML from Claude's response.")

    filepath = save_html(html_content)
    print(f"💾 HTML saved to: {filepath}")
    open_in_browser(filepath)

    return filepath

Generate with the aesthetics prompt:

BASE_SYSTEM_PROMPT = """
You are an expert frontend engineer skilled at crafting beautiful, performant frontend applications.

<tech_stack>
Use vanilla HTML, CSS, & Javascript. Use Tailwind CSS for your CSS variables.
</tech_stack>

<output>
Generate complete, self-contained HTML code for the requested frontend application. Include all CSS and JavaScript inline. 

CRITICAL: You must wrap your HTML code in triple backticks with html language identifier like this:
```html
<!DOCTYPE html>
<html>
...
</html>
```

Our parser depends on this format - do not deviate from it!
</output>
"""

USER_PROMPT = "Create a SaaS landing page for a project management tool"

# Generate with distilled aesthetics prompt
generate_html_with_claude(BASE_SYSTEM_PROMPT + "\n\n" + DISTILLED_AESTHETICS_PROMPT, USER_PROMPT)

🚀 Generating HTML...

```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Momentum — Project Management Reimagined</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@400;500;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #FF6B35;
            --primary-dark: #E85A2A;
            --secondary: #004E89;
            --accent: #FFD23F;
            --dark: #1A1A2E;
            --light: #F8F9FA;
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'DM Sans', sans-serif;
            background: var(--light);
            color: var(--dark);
            overflow-x: hidden;
        }

        h1, h2, h3, h4 {
            font-family: 'Syne', sans-serif;
            font-weight: 800;
        }

        /* Animated background */
        .hero-bg {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: linear-gradient(135deg, #004E89 0%, #1A1A2E 50%, #FF6B35 100%);
            z-index: 0;
        }

        .hero-bg::before {
            content: '';
            position: absolute;
            width: 200%;
            height: 200%;
            background: 
                radial-gradient(circle at 20% 50%, rgba(255, 107, 53, 0.3) 0%, transparent 50%),
                radial-gradient(circle at 80% 80%, rgba(255, 210, 63, 0.2) 0%, transparent 50%),
                radial-gradient(circle at 40% 20%, rgba(0, 78, 137, 0.3) 0%, transparent 50%);
            animation: float 20s ease-in-out infinite;
        }

        @keyframes float {
            0%, 100% { transform: translate(0, 0) rotate(0deg); }
            33% { transform: translate(30px, -30px) rotate(120deg); }
            66% { transform: translate(-20px, 20px) rotate(240deg); }
        }

        .mesh-gradient {
            background: 
                radial-gradient(at 27% 37%, hsla(215, 98%, 61%, 0.3) 0px, transparent 50%),
                radial-gradient(at 97% 21%, hsla(125, 98%, 72%, 0.2) 0px, transparent 50%),
                radial-gradient(at 52% 99%, hsla(354, 98%, 61%, 0.3) 0px, transparent 50%),
                radial-gradient(at 10% 29%, hsla(256, 96%, 67%, 0.2) 0px, transparent 50%);
        }

        /* Fade in animations */
        .fade-in {
            opacity: 0;
            transform: translateY(30px);
            animation: fadeInUp 0.8s ease forwards;
        }

        @keyframes fadeInUp {
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }

        .delay-1 { animation-delay: 0.1s; }
        .delay-2 { animation-delay: 0.2s; }
        .delay-3 { animation-delay: 0.3s; }
        .delay-4 { animation-delay: 0.4s; }
        .delay-5 { animation-delay: 0.5s; }
        .delay-6 { animation-delay: 0.6s; }

        /* Feature cards */
        .feature-card {
            background: white;
            border-radius: 24px;
            padding: 2rem;
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
            border: 2px solid transparent;
            position: relative;
            overflow: hidden;
        }

        .feature-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 4px;
            background: linear-gradient(90deg, var(--primary), var(--accent));
            transform: scaleX(0);
            transform-origin: left;
            transition: transform 0.4s ease;
        }

        .feature-card:hover::before {
            transform: scaleX(1);
        }

        .feature-card:hover {
            transform: translateY(-8px);
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.15);
            border-color: var(--primary);
        }

        /* CTA Button */
        .cta-button {
            background: var(--primary);
            color: white;
            padding: 1rem 2.5rem;
            border-radius: 50px;
            font-weight: 700;
            text-decoration: none;
            display: inline-block;
            transition: all 0.3s ease;
            box-shadow: 0 10px 30px rgba(255, 107, 53, 0.3);
            position: relative;
            overflow: hidden;
        }

        .cta-button::before {
            content: '';
            position: absolute;
            top: 50%;
            left: 50%;
            width: 0;
            height: 0;
            border-radius: 50%;
            background: rgba(255, 255, 255, 0.2);
            transform: translate(-50%, -50%);
            transition: width 0.6s, height 0.6s;
        }

        .cta-button:hover::before {
            width: 300px;
            height: 300px;
        }

        .cta-button:hover {
            transform: translateY(-2px);
            box-shadow: 0 15px 40px rgba(255, 107, 53, 0.4);
        }

        .cta-button span {
            position: relative;
            z-index: 1;
        }

        /* Stats counter animation */
        .stat-number {
            font-size: 3rem;
            font-weight: 800;
            background: linear-gradient(135deg, var(--primary), var(--accent));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }

        /* Testimonial cards */
        .testimonial {
            background: white;
            padding: 2rem;
            border-radius: 20px;
            box-shadow: 0 10px 40px rgba(0, 0, 0, 0.08);
            transition: transform 0.3s ease;
        }

        .testimonial:hover {
            transform: scale(1.02);
        }

        /* Icon styles */
        .icon-circle {
            width: 60px;
            height: 60px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 24px;
            margin-bottom: 1rem;
        }

        /* Navbar scroll effect */
        .navbar {
            transition: all 0.3s ease;
        }

        .navbar.scrolled {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
            box-shadow: 0 2px 20px rgba(0, 0, 0, 0.1);
        }

        /* Pricing cards */
        .pricing-card {
            background: white;
            border-radius: 24px;
            padding: 3rem 2rem;
            transition: all 0.4s ease;
            border: 2px solid #e5e7eb;
        }

        .pricing-card.featured {
            border-color: var(--primary);
            transform: scale(1.05);
            box-shadow: 0 20px 60px rgba(255, 107, 53, 0.2);
        }

        .pricing-card:hover {
            transform: translateY(-10px) scale(1.02);
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.15);
        }

        /* Dashboard mockup */
        .dashboard-mockup {
            background: white;
            border-radius: 20px;
            box-shadow: 0 30px 80px rgba(0, 0, 0, 0.2);
            padding: 1rem;
            position: relative;
            transform: perspective(1000px) rotateY(-5deg) rotateX(5deg);
            transition: transform 0.5s ease;
        }

        .dashboard-mockup:hover {
            transform: perspective(1000px) rotateY(0deg) rotateX(0deg);
        }

        .mockup-header {
            display: flex;
            gap: 8px;
            margin-bottom: 1rem;
        }

        .mockup-dot {
            width: 12px;
            height: 12px;
            border-radius: 50%;
        }

        .mockup-content {
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            border-radius: 12px;
            height: 400px;
            position: relative;
            overflow: hidden;
        }

        .mockup-element {
            position: absolute;
            background: white;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
        }
    </style>
</head>
<body>
    <!-- Navigation -->
    <nav class="navbar fixed w-full top-0 z-50 py-4 px-8">
        <div class="max-w-7xl mx-auto flex items-center justify-between">
            <div class="flex items-center gap-2">
                <div class="w-10 h-10 rounded-full bg-gradient-to-br from-orange-500 to-yellow-400 flex items-center justify-center">
                    <span class="text-white font-bold text-xl">M</span>
                </div>
                <span class="text-2xl font-bold">Momentum</span>
            </div>
            <div class="hidden md:flex items-center gap-8">
                <a href="#features" class="text-gray-700 hover:text-orange-500 transition font-medium">Features</a>
                <a href="#pricing" class="text-gray-700 hover:text-orange-500 transition font-medium">Pricing</a>
                <a href="#testimonials" class="text-gray-700 hover:text-orange-500 transition font-medium">Testimonials</a>
                <a href="#" class="text-gray-700 hover:text-orange-500 transition font-medium">Login</a>
                <a href="#" class="cta-button"><span>Start Free Trial</span></a>
            </div>
            <button class="md:hidden text-gray-700">
                <svg width="24" height="24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <line x1="3" y1="12" x2="21" y2="12"></line>
                    <line x1="3" y1="6" x2="21" y2="6"></line>
                    <line x1="3" y1="18" x2="21" y2="18"></line>
                </svg>
            </button>
        </div>
    </nav>

    <!-- Hero Section -->
    <section class="relative min-h-screen flex items-center justify-center overflow-hidden pt-20">
        <div class="hero-bg"></div>
        
        <div class="relative z-10 max-w-7xl mx-auto px-8 py-20">
            <div class="grid md:grid-cols-2 gap-12 items-center">
                <div class="text-white">
                    <h1 class="text-6xl md:text-7xl leading-tight mb-6 fade-in">
                        Build momentum.<br/>
                        <span class="text-yellow-300">Ship faster.</span>
                    </h1>
                    <p class="text-xl mb-8 text-gray-200 fade-in delay-1">
                        The project management tool that adapts to your team's rhythm. Stop managing tasks. Start building momentum.
                    </p>
                    <div class="flex flex-wrap gap-4 fade-in delay-2">
                        <a href="#" class="cta-button"><span>Get Started Free</span></a>
                        <a href="#" class="bg-white/10 backdrop-blur text-white px-8 py-4 rounded-full font-bold hover:bg-white/20 transition inline-block">
                            Watch Demo
                        </a>
                    </div>
                    <div class="flex items-center gap-8 mt-12 fade-in delay-3">
                        <div>
                            <div class="stat-number text-white">50k+</div>
                            <div class="text-gray-300">Active Teams</div>
                        </div>
                        <div>
                            <div class="stat-number text-white">4.9</div>
                            <div class="text-gray-300">Average Rating</div>
                        </div>
                        <div>
                            <div class="stat-number text-white">99%</div>
                            <div class="text-gray-300">Uptime</div>
                        </div>
                    </div>
                </div>
                
                <div class="fade-in delay-4">
                    <div class="dashboard-mockup">
                        <div class="mockup-header">
                            <div class="mockup-dot bg-red-500"></div>
                            <div class="mockup-dot bg-yellow-400"></div>
                            <div class="mockup-dot bg-green-500"></div>
                        </div>
                        <div class="mockup-content">
                            <div class="mockup-element" style="top: 20px; left: 20px; width: 200px; height: 60px;"></div>
                            <div class="mockup-element" style="top: 100px; left: 20px; width: 150px; height: 100px;"></div>
                            <div class="mockup-element" style="top: 100px; right: 20px; width: 150px; height: 100px;"></div>
                            <div class="mockup-element" style="top: 220px; left: 20px; width: 320px; height: 80px;"></div>
                            <div class="mockup-element" style="bottom: 20px; right: 20px; width: 100px; height: 60px; background: linear-gradient(135deg, #FF6B35, #FFD23F);"></div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </section>

    <!-- Features Section -->
    <section id="features" class="py-32 px-8 bg-white">
        <div class="max-w-7xl mx-auto">
            <div class="text-center mb-20">
                <h2 class="text-5xl md:text-6xl font-bold mb-6">Everything you need.<br/>Nothing you don't.</h2>
                <p class="text-xl text-gray-600 max-w-2xl mx-auto">Powerful features that don't get in your way. Built for teams who want to focus on work, not tools.</p>
            </div>

            <div class="grid md:grid-cols-3 gap-8">
                <div class="feature-card">
                    <div class="icon-circle bg-orange-100 text-orange-500">⚡</div>
                    <h3 class="text-2xl font-bold mb-4">Lightning Fast</h3>
                    <p class="text-gray-600">Native performance across all devices. No lag, no loading spinners. Just instant productivity.</p>
                </div>

                <div class="feature-card">
                    <div class="icon-circle bg-blue-100 text-blue-500">🎯</div>
                    <h3 class="text-2xl font-bold mb-4">Smart Workflows</h3>
                    <p class="text-gray-600">AI-powered automation that learns from your team's patterns and suggests optimizations.</p>
                </div>

                <div class="feature-card">
                    <div class="icon-circle bg-purple-100 text-purple-500">🔗</div>
                    <h3 class="text-2xl font-bold mb-4">Seamless Integration</h3>
                    <p class="text-gray-600">Connect with 1000+ tools your team already uses. Slack, GitHub, Figma, and more.</p>
                </div>

                <div class="feature-card">
                    <div class="icon-circle bg-green-100 text-green-500">📊</div>
                    <h3 class="text-2xl font-bold mb-4">Real-time Analytics</h3>
                    <p class="text-gray-600">Visualize progress with beautiful dashboards that update in real-time as work happens.</p>
                </div>

                <div class="feature-card">
                    <div class="icon-circle bg-pink-100 text-pink-500">🎨</div>
                    <h3 class="text-2xl font-bold mb-4">Customizable Views</h3>
                    <p class="text-gray-600">Board, list, timeline, calendar - switch between views instantly. See work your way.</p>
                </div>

                <div class="feature-card">
                    <div class="icon-circle bg-yellow-100 text-yellow-600">🔒</div>
                    <h3 class="text-2xl font-bold mb-4">Enterprise Security</h3>
                    <p class="text-gray-600">SOC 2 Type II certified. Your data encrypted at rest and in transit. Always.</p>
                </div>
            </div>
        </div>
    </section>

    <!-- Social Proof / Stats -->
    <section class="py-24 px-8 bg-gradient-to-br from-gray-900 to-gray-800 text-white">
        <div class="max-w-7xl mx-auto">
            <div class="text-center mb-16">
                <h2 class="text-4xl md:text-5xl font-bold mb-4">Trusted by teams worldwide</h2>
                <p class="text-xl text-gray-300">Join thousands of companies building better products</p>
            </div>
            
            <div class="grid grid-cols-2 md:grid-cols-4 gap-8 mb-16">
                <div class="text-center">
                    <div class="text-5xl font-bold mb-2 text-yellow-400">2.5M+</div>
                    <div class="text-gray-400">Projects Created</div>
                </div>
                <div class="text-center">
                    <div class="text-5xl font-bold mb-2 text-yellow-400">50K+</div>
                    <div class="text-gray-400">Active Teams</div>
                </div>
                <div class="text-center">
                    <div class="text-5xl font-bold mb-2 text-yellow-400">150+</div>
                    <div class="text-gray-400">Countries</div>
                </div>
                <div class="text-center">
                    <div class="text-5xl font-bold mb-2 text-yellow-400">99.9%</div>
                    <div class="text-gray-400">Uptime SLA</div>
                </div>
            </div>

            <div class="flex flex-wrap justify-center items-center gap-12 opacity-60">
                <div class="text-3xl font-bold">Stripe</div>
                <div class="text-3xl font-bold">Notion</div>
                <div class="text-3xl font-bold">Figma</div>
                <div class="text-3xl font-bold">Webflow</div>
                <div class="text-3xl font-bold">Linear</div>
            </div>
        </div>
    </section>

    <!-- Testimonials -->
    <section id="testimonials" class="py-32 px-8 bg-gray-50">
        <div class="max-w-7xl mx-auto">
            <div class="text-center mb-20">
                <h2 class="text-5xl font-bold mb-6">Loved by teams everywhere</h2>
                <p class="text-xl text-gray-600">Don't just take our word for it</p>
            </div>

            <div class="grid md:grid-cols-3 gap-8">
                <div class="testimonial">
                    <div class="flex items-center gap-1 mb-4">
                        <span class="text-yellow-400 text-xl">★★★★★</span>
                    </div>
                    <p class="text-gray-700 mb-6">"Momentum completely changed how our team works. We shipped our last feature 40% faster than usual. The automation is brilliant."</p>
                    <div class="flex items-center gap-3">
                        <div class="w-12 h-12 rounded-full bg-gradient-to-br from-purple-400 to-pink-400"></div>
                        <div>
                            <div class="font-bold">Sarah Chen</div>
                            <div class="text-sm text-gray-500">Head of Product, TechCorp</div>
                        </div>
                    </div>
                </div>

                <div class="testimonial">
                    <div class="flex items-center gap-1 mb-4">
                        <span class="text-yellow-400 text-xl">★★★★★</span>
                    </div>
                    <p class="text-gray-700 mb-6">"Finally, a project management tool that doesn't feel like homework. Our team adoption rate was 100% in the first week."</p>
                    <div class="flex items-center gap-3">
                        <div class="w-12 h-12 rounded-full bg-gradient-to-br from-blue-400 to-cyan-400"></div>
                        <div>
                            <div class="font-bold">Marcus Rodriguez</div>
                            <div class="text-sm text-gray-500">Engineering Manager, StartupXYZ</div>
                        </div>
                    </div>
                </div>

                <div class="testimonial">
                    <div class="flex items-center gap-1 mb-4">
                        <span class="text-yellow-400 text-xl">★★★★★</span>
                    </div>
                    <p class="text-gray-700 mb-6">"The real-time collaboration features are next level. It's like Google Docs, but for project management. Game changer."</p>
                    <div class="flex items-center gap-3">
                        <div class="w-12 h-12 rounded-full bg-gradient-to-br from-green-400 to-emerald-400"></div>
                        <div>
                            <div class="font-bold">Aisha Patel</div>
                            <div class="text-sm text-gray-500">Design Lead, CreativeStudio</div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </section>

    <!-- Pricing -->
    <section id="pricing" class="py-32 px-8 bg-white">
        <div class="max-w-7xl mx-auto">
            <div class="text-center mb-20">
                <h2 class="text-5xl md:text-6xl font-bold mb-6">Simple, transparent pricing</h2>
                <p class="text-xl text-gray-600">No hidden fees. Cancel anytime. Start with a 14-day free trial.</p>
            </div>

            <div class="grid md:grid-cols-3 gap-8 max-w-6xl mx-auto">
                <div class="pricing-card">
                    <div class="text-sm font-bold text-gray-500 mb-2">STARTER</div>
                    <div class="mb-6">
                        <span class="text-5xl font-bold">$12</span>
                        <span class="text-gray-500">/user/month</span>
                    </div>
                    <ul class="space-y-4 mb-8">
                        <li class="flex items-center gap-2">
                            <span class="text-green-500">✓</span>
                            <span>Up to 10 team members</span>
                        </li>
                        <li class="flex items-center gap-2">
                            <span class="text-green-500">✓</span>
                            <span>Unlimited projects</span>
                        </li>
                        <li class="flex items-center gap-2">
                            <span class="text-green-500">✓</span>
                            <span>Basic integrations</span>
                        </li>
                        <li class="flex items-center gap-2">
                            <span class="text-green-500">✓</span>
                            <span>5GB storage</span>
                        </li>
                    </ul>
                    <a href="#" class="block text-center bg-gray-900 text-white py-3 rounded-full font-bold hover:bg-gray-800 transition">
                        Start Free Trial
                    </a>
                </div>

                <div class="pricing-card featured">
                    <div class="text-sm font-bold text-orange-500 mb-2">PROFESSIONAL</div>
                    <div class="mb-6">
                        <span class="text-5xl font-bold">$29</span>
                        <span class="text-gray-500">/user/month</span>
                    </div>
                    <ul class="space-y-4 mb-8">
                        <li class="flex items-center gap-2">
                            <span class="text-green-500">✓</span>
                            <span>Unlimited team members</span>
                        </li>
                        <li class="flex items-center gap-2">
                            <span class="text-green-500">✓</span>
                            <span>Advanced automation</span>
                        </li>
                        <li class="flex items-center gap-2">
                            <span class="text-green-500">✓</span>
                            <span>All integrations</span>
                        </li>
                        <li class="flex items-center gap-2">
                            <span class="text-green-500">✓</span>
                            <span>100GB storage</span>
                        </li>
                        <li class="flex items-center gap-2">
                            <span class="text-green-500">✓</span>
                            <span>Priority support</span>
                        </li>
                    </ul>
                    <a href="#" class="block text-center bg-orange-500 text-white py-3 rounded-full font-bold hover:bg-orange-600 transition">
                        Start Free Trial
                    </a>
                </div>

                <div class="pricing-card">
                    <div class="text-sm font-bold text-gray-500 mb-2">ENTERPRISE</div>
                    <div class="mb-6">
                        <span class="text-5xl font-bold">Custom</span>
                    </div>
                    <ul class="space-y-4 mb-8">
                        <li class="flex items-center gap-2">
                            <span class="text-green-500">✓</span>
                            <span>Everything in Pro</span>
                        </li>
                        <li class="flex items-center gap-2">
                            <span class="text-green-500">✓</span>
                            <span>Advanced security</span>
                        </li>
                        <li class="flex items-center gap-2">
                            <span class="text-green-500">✓</span>
                            <span>Custom integrations</span>
                        </li>
                        <li class="flex items-center gap-2">
                            <span class="text-green-500">✓</span>
                            <span>Unlimited storage</span>
                        </li>
                        <li class="flex items-center gap-2">
                            <span class="text-green-500">✓</span>
                            <span>Dedicated support</span>
                        </li>
                    </ul>
                    <a href="#" class="block text-center bg-gray-900 text-white py-3 rounded-full font-bold hover:bg-gray-800 transition">
                        Contact Sales
                    </a>
                </div>
            </div>
        </div>
    </section>

    <!-- CTA Section -->
    <section class="py-32 px-8 bg-gradient-to-br from-orange-500 to-yellow-400 text-white">
        <div class="max-w-4xl mx-auto text-center">
            <h2 class="text-5xl md:text-6xl font-bold mb-6">Ready to build momentum?</h2>
            <p class="text-2xl mb-12 text-white/90">Join 50,000+ teams shipping faster with Momentum</p>
            <div class="flex flex-wrap gap-4 justify-center">
                <a href="#" class="bg-white text-orange-500 px-10 py-5 rounded-full font-bold text-lg hover:bg-gray-100 transition inline-block">
                    Start Free Trial
                </a>
                <a href="#" class="bg-white/10 backdrop-blur text-white px-10 py-5 rounded-full font-bold text-lg hover:bg-white/20 transition inline-block">
                    Schedule Demo
                </a>
            </div>
            <p class="mt-8 text-white/80">No credit card required • 14-day free trial • Cancel anytime</p>
        </div>
    </section>

    <!-- Footer -->
    <footer class="bg-gray-900 text-white py-16 px-8">
        <div class="max-w-7xl mx-auto">
            <div class="grid md:grid-cols-4 gap-12 mb-12">
                <div>
                    <div class="flex items-center gap-2 mb-4">
                        <div class="w-10 h-10 rounded-full bg-gradient-to-br from-orange-500 to-yellow-400 flex items-center justify-center">
                            <span class="text-white font-bold text-xl">M</span>
                        </div>
                        <span class="text-2xl font-bold">Momentum</span>
                    </div>
                    <p class="text-gray-400">Building momentum for teams that ship.</p>
                </div>
                
                <div>
                    <h4 class="font-bold mb-4">Product</h4>
                    <ul class="space-y-2 text-gray-400">
                        <li><a href="#" class="hover:text-white transition">Features</a></li>
                        <li><a href="#" class="hover:text-white transition">Pricing</a></li>
                        <li><a href="#" class="hover:text-white transition">Integrations</a></li>
                        <li><a href="#" class="hover:text-white transition">Changelog</a></li>
                    </ul>
                </div>
                
                <div>
                    <h4 class="font-bold mb-4">Company</h4>
                    <ul class="space-y-2 text-gray-400">
                        <li><a href="#" class="hover:text-white transition">About</a></li>
                        <li><a href="#" class="hover:text-white transition">Blog</a></li>
                        <li><a href="#" class="hover:text-white transition">Careers</a></li>
                        <li><a href="#" class="hover:text-white transition">Contact</a></li>
                    </ul>
                </div>
                
                <div>
                    <h4 class="font-bold mb-4">Legal</h4>
                    <ul class="space-y-2 text-gray-400">
                        <li><a href="#" class="hover:text-white transition">Privacy</a></li>
                        <li><a href="#" class="hover:text-white transition">Terms</a></li>
                        <li><a href="#" class="hover:text-white transition">Security</a></li>
                        <li><a href="#" class="hover:text-white transition">GDPR</a></li>
                    </ul>
                </div>
            </div>
            
            <div class="border-t border-gray-800 pt-8 flex flex-col md:flex-row justify-between items-center gap-4">
                <p class="text-gray-400">© 2024 Momentum. All rights reserved.</p>
                <div class="flex gap-6">
                    <a href="#" class="text-gray-400 hover:text-white transition">Twitter</a>
                    <a href="#" class="text-gray-400 hover:text-white transition">LinkedIn</a>
                    <a href="#" class="text-gray-400 hover:text-white transition">GitHub</a>
                </div>
            </div>
        </div>
    </footer>

    <script>
        // Navbar scroll effect
        const navbar = document.querySelector('.navbar');
        window.addEventListener('scroll', () => {
            if (window.scrollY > 50) {
                navbar.classList.add('scrolled');
            } else {
                navbar.classList.remove('scrolled');
            }
        });

        // Smooth scroll for anchor links
        document.querySelectorAll('a[href^="#"]').forEach(anchor => {
            anchor.addEventListener('click', function (e) {
                e.preventDefault();
                const target = document.querySelector(this.getAttribute('href'));
                if (target) {
                    target.scrollIntoView({
                        behavior: 'smooth',
                        block: 'start'
                    });
                }
            });
        });

        // Intersection Observer for fade-in animations
        const observerOptions = {
            threshold: 0.1,
            rootMargin: '0px 0px -50px 0px'
        };

        const observer = new IntersectionObserver((entries) => {
            entries.forEach(entry => {
                if (entry.isIntersecting) {
                    entry.target.style.opacity = '1';
                    entry.target.style.transform = 'translateY(0)';
                }
            });
        }, observerOptions);

        document.querySelectorAll('.feature-card, .testimonial, .pricing-card').forEach(el => {
            el.style.opacity = '0';
            el.style.transform = 'translateY(30px)';
            el.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
            observer.observe(el);
        });
    </script>
</body>
</html>
```

✅ Complete in 98.2s

💾 HTML saved to: html_outputs/20251021_101010.html
🌐 Opened in browser: html_outputs/20251021_101010.html

'html_outputs/20251021_101010.html'

## Isolated Prompting

The full aesthetics prompt works well for general use, but sometimes you want targeted control. You can isolate specific dimensions (typography, color, motion) or lock in a particular theme. This gives you faster generation times and more predictable outputs.

### Example 1: Typography Only

Isolate a single design dimension when you want to improve one aspect without changing others:

TYPOGRAPHY_PROMPT = """
<use_interesting_fonts>
Typography instantly signals quality. Avoid using boring, generic fonts.

**Never use:** Inter, Roboto, Open Sans, Lato, default system fonts

**Impact choices:**
- Code aesthetic: JetBrains Mono, Fira Code, Space Grotesk
- Editorial: Playfair Display, Crimson Pro, Fraunces
- Startup: Clash Display, Satoshi, Cabinet Grotesk
- Technical: IBM Plex family, Source Sans 3
- Distinctive: Bricolage Grotesque, Obviously, Newsreader

**Pairing principle:** High contrast = interesting. Display + monospace, serif + geometric sans, variable font across weights.

**Use extremes:** 100/200 weight vs 800/900, not 400 vs 600. Size jumps of 3x+, not 1.5x.

Pick one distinctive font, use it decisively. Load from Google Fonts. State your choice before coding.
</use_interesting_fonts>
"""

# Generate with typography-only guidance
generate_html_with_claude(BASE_SYSTEM_PROMPT + "\n\n" + TYPOGRAPHY_PROMPT, USER_PROMPT)

### Example 2: Theme Constraint

Lock in a specific aesthetic when you want consistent theming across generations:

SOLARPUNK_THEME_PROMPT = """
<always_use_solarpunk_theme>
Always design with Solarpunk aesthetic:
- Warm, optimistic color palettes (greens, golds, earth tones)
- Organic shapes mixed with technical elements
- Nature-inspired patterns and textures
- Bright, hopeful atmosphere
- Retro-futuristic typography
</always_use_solarpunk_theme>
"""

# Generate with theme constraint
generate_html_with_claude(
    BASE_SYSTEM_PROMPT + "\n\n" + SOLARPUNK_THEME_PROMPT,
    "Create a dashboard for renewable energy monitoring",
)

## Summary

Claude has strong design capabilities but defaults to safe, generic choices. The techniques in this guide - targeting specific design dimensions, referencing concrete inspirations, and explicitly avoiding common defaults - reliably produce more distinctive output. The full aesthetics prompt works well as a baseline. For more control, use isolated prompts to focus on individual aspects or lock in specific themes across multiple generations.