#!/usr/bin/env python3
"""
Create a PowerPoint presentation on LLMs as a more efficient way of gathering information.
Based on Nicolas Bustamante's "LLM Context Tax" article and related research.
"""

from pptx import Presentation
from pptx.util import Inches, Pt


def add_title_slide(prs, title, subtitle=""):
    """Add a title slide."""
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    if subtitle and slide.placeholders[1]:
        slide.placeholders[1].text = subtitle
    return slide


def add_content_slide(prs, title, bullet_points):
    """Add a content slide with title and bullet points."""
    slide_layout = prs.slide_layouts[1]  # Title and Content
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    body = slide.placeholders[1]
    tf = body.text_frame
    tf.clear()
    for i, point in enumerate(bullet_points):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = point
        p.level = 0
        p.space_after = Pt(12)
    return slide


def add_section_slide(prs, title):
    """Add a section header slide."""
    slide_layout = prs.slide_layouts[5]  # Section header
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    return slide


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Slide 1: Title
    add_title_slide(
        prs,
        "LLMs: A More Efficient Way of Gathering Information",
        "Based on Nicolas Bustamante's research on the LLM Context Tax and information efficiency"
    )

    # Slide 2: The Problem - Context Tax
    add_content_slide(
        prs,
        "The LLM Context Tax: Why Efficiency Matters",
        [
            "The 'Context Tax' = cost penalty from filling context with useless tokens",
            "Triple penalty: Higher costs + Slower responses + Degraded performance",
            "Context rot: Agents become confused in accumulated noise",
            "Past 32K tokens, most models show sharp performance degradation",
        ]
    )

    # Slide 3: Cost Structure
    add_content_slide(
        prs,
        "Understanding LLM Cost Structure",
        [
            "Cached inputs cost ~10x less than uncached inputs (e.g., Claude Opus)",
            "Output tokens cost ~5x more than uncached input tokens",
            "The cheapest token is one never sent—application-level caching wins",
            "Key insight: Optimize for cache hits to achieve dramatic cost savings",
        ]
    )

    # Slide 4: KV Cache Optimization
    add_content_slide(
        prs,
        "KV Cache: The Key to 10x Efficiency Gains",
        [
            "Stable prefixes (identical prompt beginnings) enable cached computations",
            "Move dynamic content (timestamps, user IDs) to the END of prompts",
            "Append-only context: Mutating context destroys cache hit rates",
            "Precise tool design can reduce token consumption by 10x",
        ]
    )

    # Slide 5: Context Engineering
    add_content_slide(
        prs,
        "Context Engineering Best Practices",
        [
            "Place static content first—system prompts, instructions, documentation",
            "Reserve the end for dynamic content—queries, timestamps, session data",
            "Design tools that return minimal, precise responses",
            "Output token budgeting: Minimize generation since output costs 5x more",
        ]
    )

    # Slide 6: Measuring Efficiency
    add_content_slide(
        prs,
        "Measuring Information Efficiency",
        [
            "OckBench: Models with comparable accuracy differ significantly in token use",
            "Efficiency is a neglected but critical differentiation axis",
            "Each task has intrinsic minimal token requirements",
            "Current strategies operate far from theoretical limits—room for improvement",
        ]
    )

    # Slide 7: LLM vs Human Compression
    add_content_slide(
        prs,
        "LLM vs. Human Information Compression",
        [
            "LLMs compress aggressively for optimal information-theoretic efficiency",
            "Humans maintain 'inefficient' representations preserving contextual nuance",
            "Tradeoff: Compression vs. semantic richness",
            "Smaller encoder models can outperform larger decoders on human-aligned tasks",
        ]
    )

    # Slide 8: Adaptive Reasoning
    add_content_slide(
        prs,
        "Adaptive Reasoning: Match Effort to Difficulty",
        [
            "Thinking models often 'overthink' simple problems—wasting computation",
            "Use simpler models for easy problems, reasoning models for harder ones",
            "Adaptive approaches: 21% token reduction while maintaining accuracy",
            "Right-size your model to the task complexity",
        ]
    )

    # Slide 9: Key Takeaways
    add_content_slide(
        prs,
        "Key Takeaways: LLMs as Efficient Information Gatherers",
        [
            "Optimize for KV cache hits—structure prompts for stable prefixes",
            "Context engineering can yield 10x token reduction",
            "Output tokens are expensive—minimize generation, maximize caching",
            "Efficiency deserves equal attention to accuracy in production systems",
        ]
    )

    # Slide 10: Conclusion
    add_title_slide(
        prs,
        "LLMs: Efficient by Design",
        "With proper context engineering, LLMs become dramatically more efficient at gathering and processing information—reducing costs, improving speed, and maintaining performance."
    )

    # Save
    output_path = "/workspace/LLM_Information_Efficiency.pptx"
    prs.save(output_path)
    print(f"Presentation saved to {output_path}")


if __name__ == "__main__":
    main()
