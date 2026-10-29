import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_light_deck(filename="Phase_4_Final_Project_Presentation_Team65_Veltrox.pptx"):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Clean, High-Contrast Light Executive Palette
    c_bg = RGBColor(248, 250, 252)          # Slate 50 (Soft Off-White)
    c_card = RGBColor(255, 255, 255)        # Pure White
    c_card_border = RGBColor(226, 232, 240) # Slate 200 (Crisp Border)
    c_card_border_strong = RGBColor(203, 213, 225) # Slate 300
    c_card_subtle = RGBColor(241, 245, 249) # Slate 100
    
    # Core Brand & Category Accents
    c_teal = RGBColor(13, 148, 136)         # Teal 600
    c_teal_dark = RGBColor(15, 118, 110)    # Teal 700
    c_teal_light = RGBColor(204, 251, 241)  # Teal 100
    
    c_blue = RGBColor(37, 99, 235)          # Blue 600
    c_blue_dark = RGBColor(29, 78, 216)     # Blue 700
    c_blue_light = RGBColor(219, 234, 254)  # Blue 100
    
    c_emerald = RGBColor(5, 150, 105)       # Emerald 600
    c_emerald_light = RGBColor(209, 250, 229) # Emerald 100
    
    c_amber = RGBColor(217, 119, 6)         # Amber 600
    c_amber_light = RGBColor(254, 243, 199) # Amber 100

    c_purple = RGBColor(124, 58, 237)       # Purple 600
    c_purple_light = RGBColor(243, 232, 255) # Purple 100

    # Typography
    c_text_primary = RGBColor(15, 23, 42)   # Slate 900 (High contrast dark slate)
    c_text_secondary = RGBColor(51, 65, 85) # Slate 700 (Clear readable body)
    c_text_muted = RGBColor(100, 116, 139)  # Slate 500
    c_white = RGBColor(255, 255, 255)

    # Pre-instantiate all 14 slides so we can cross-reference them for interactive links!
    slides = [prs.slides.add_slide(blank_layout) for _ in range(14)]
    (s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14) = slides

    def add_slide_background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = c_bg
        bg.line.fill.background()
        return bg

    def add_interactive_header(slide, current_idx, title_text, category_text="PHASE 4: FINAL PROJECT PRESENTATION"):
        # Top background bar
        hdr_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.15))
        hdr_bar.fill.solid()
        hdr_bar.fill.fore_color.rgb = c_white
        hdr_bar.line.color.rgb = c_card_border
        hdr_bar.line.width = Pt(1)

        # Accent top line
        top_accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.06))
        top_accent.fill.solid()
        top_accent.fill.fore_color.rgb = c_teal
        top_accent.line.fill.background()

        # Breadcrumb & Slide Title
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.12), Inches(6.5), Inches(0.95))
        p_cat = tb.text_frame.paragraphs[0]
        r_c1 = p_cat.add_run()
        r_c1.text = "TEAM 65 (VELTROX) • HACKMITTEN 2026 • "
        r_c1.font.bold = True
        r_c1.font.size = Pt(8.5)
        r_c1.font.color.rgb = c_teal_dark

        r_c2 = p_cat.add_run()
        r_c2.text = category_text
        r_c2.font.bold = False
        r_c2.font.size = Pt(8.5)
        r_c2.font.color.rgb = c_text_muted

        p_title = tb.text_frame.add_paragraph()
        r_t = p_title.add_run()
        r_t.text = title_text
        r_t.font.bold = True
        r_t.font.size = Pt(16.5)
        r_t.font.color.rgb = c_text_primary

        # Interactive Navigation Buttons Bar (Clickable tabs to jump across slides!)
        nav_defs = [
            ("Overview", s2, 1),
            ("Problem", s3, 2),
            ("Solution", s4, 3),
            ("Methodology", s5, 4),
            ("Architecture", s6, 5),
            ("Stack", s7, 6),
            ("Features", s8, 7),
            ("Testing", s9, 8),
            ("Screenshots", s10, 9),
            ("Challenges", s12, 11),
            ("Roadmap", s13, 12),
        ]

        nav_x_start = Inches(7.1)
        nav_btn_w = Inches(0.55)
        nav_gap = Inches(0.04)

        for i, (n_lbl, target_slide, target_idx) in enumerate(nav_defs):
            x = nav_x_start + i * (nav_btn_w + nav_gap)
            is_active = (current_idx == target_idx)
            
            n_btn = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(0.25), nav_btn_w, Inches(0.3))
            n_btn.fill.solid()
            n_btn.fill.fore_color.rgb = c_teal_light if is_active else c_card_subtle
            n_btn.line.color.rgb = c_teal if is_active else c_card_border
            n_btn.line.width = Pt(1)

            tf = n_btn.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = n_lbl
            r.font.bold = is_active
            r.font.size = Pt(7.5)
            r.font.color.rgb = c_teal_dark if is_active else c_text_secondary
            r.hyperlink.target_slide = target_slide

        # Live Demo & GitHub External Links
        b_demo = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.45), Inches(0.68), Inches(1.05), Inches(0.34))
        b_demo.fill.solid()
        b_demo.fill.fore_color.rgb = c_teal
        b_demo.line.color.rgb = c_teal_dark
        tf_d = b_demo.text_frame
        tf_d.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_d = tf_d.paragraphs[0]
        p_d.alignment = PP_ALIGN.CENTER
        r_d = p_d.add_run()
        r_d.text = "⚡ Live Demo"
        r_d.font.bold = True
        r_d.font.size = Pt(8.5)
        r_d.font.color.rgb = c_white
        r_d.hyperlink.address = "https://chatvia-app-k7o7.onrender.com/"

        b_gh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(12.55), Inches(0.68), Inches(0.7), Inches(0.34))
        b_gh.fill.solid()
        b_gh.fill.fore_color.rgb = c_card_subtle
        b_gh.line.color.rgb = c_card_border_strong
        tf_g = b_gh.text_frame
        tf_g.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_g = tf_g.paragraphs[0]
        p_g.alignment = PP_ALIGN.CENTER
        r_g = p_g.add_run()
        r_g.text = "GitHub"
        r_g.font.bold = True
        r_g.font.size = Pt(8.5)
        r_g.font.color.rgb = c_text_primary
        r_g.hyperlink.address = "https://github.com/Kishan-Kumar-Hv/hmit"

    def add_card(slide, left, top, width, height, title="", title_color=c_teal_dark, border_color=c_card_border, bg_color=c_card, badge_text="", badge_bg=c_teal_light, badge_color=c_teal_dark):
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        box.fill.solid()
        box.fill.fore_color.rgb = bg_color
        box.line.color.rgb = border_color
        box.line.width = Pt(1)

        if title:
            tb = slide.shapes.add_textbox(left + Inches(0.18), top + Inches(0.12), width - Inches(0.36), Inches(0.4))
            p = tb.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = title
            r.font.bold = True
            r.font.size = Pt(11)
            r.font.color.rgb = title_color

            if badge_text:
                r_sp = p.add_run()
                r_sp.text = "   "
                r_b = p.add_run()
                r_b.text = f" {badge_text} "
                r_b.font.bold = True
                r_b.font.size = Pt(7.5)
                r_b.font.color.rgb = badge_color
        return box

    def add_jump_button(slide, left, top, width, height, text, target_slide, bg_color=c_card_subtle, border_color=c_card_border_strong, text_color=c_text_primary):
        btn = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        btn.fill.solid()
        btn.fill.fore_color.rgb = bg_color
        btn.line.color.rgb = border_color
        btn.line.width = Pt(1)
        tf = btn.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = text
        r.font.bold = True
        r.font.size = Pt(8)
        r.font.color.rgb = text_color
        r.hyperlink.target_slide = target_slide
        return btn

    # Initialize backgrounds for all slides
    for s in slides:
        add_slide_background(s)

    # =========================================================================
    # SLIDE 1: TITLE & TEAM DETAILS (Light Executive Theme)
    # =========================================================================
    # Top Accent Strip
    strip = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.1))
    strip.fill.solid()
    strip.fill.fore_color.rgb = c_teal
    strip.line.fill.background()

    # Track Badge
    badge1 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.65), Inches(3.9), Inches(0.38))
    badge1.fill.solid()
    badge1.fill.fore_color.rgb = c_teal_light
    badge1.line.color.rgb = c_teal
    badge1.line.width = Pt(1)
    tf_b1 = badge1.text_frame
    tf_b1.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_b1 = tf_b1.paragraphs[0]
    p_b1.alignment = PP_ALIGN.CENTER
    r_b1 = p_b1.add_run()
    r_b1.text = "HACKMITTEN 2026 • CYBERSECURITY & AI TRACK"
    r_b1.font.bold = True
    r_b1.font.size = Pt(9)
    r_b1.font.color.rgb = c_teal_dark

    # Project Title Box
    t_title = s1.shapes.add_textbox(Inches(0.8), Inches(1.15), Inches(11.733), Inches(1.6))
    p_t = t_title.text_frame.paragraphs[0]
    r_t = p_t.add_run()
    r_t.text = "CHATVIA"
    r_t.font.bold = True
    r_t.font.size = Pt(40)
    r_t.font.color.rgb = c_text_primary

    p_sub = t_title.text_frame.add_paragraph()
    r_sub = p_sub.add_run()
    r_sub.text = "Next-Generation Confidential Messaging with AI Biometrics & Client-Side Steganography"
    r_sub.font.bold = True
    r_sub.font.size = Pt(16)
    r_sub.font.color.rgb = c_teal_dark

    p_desc = t_title.text_frame.add_paragraph()
    r_desc = p_desc.add_run()
    r_desc.text = "Eliminating the physical 'last-inch' presentation vulnerability through continuous facial verification, zero-knowledge encryption, and ephemeral forensic self-healing."
    r_desc.font.size = Pt(11)
    r_desc.font.color.rgb = c_text_secondary

    # Team 65 / Veltrox Card Container
    add_card(s1, Inches(0.8), Inches(2.95), Inches(11.733), Inches(3.6), title="TEAM 65 — VELTROX | MEMBER ROLES & CONTRIBUTIONS", title_color=c_text_primary)

    members = [
        ("Kishan Kumar H V", "Team Lead & Cloud Architect", "Production PaaS deployment on Render, Express 5 reverse proxy, Supabase SSL pooler engineering, system testing, automated build pipelines, and quality assurance.", c_teal, c_teal_light, c_teal_dark),
        ("Santhosh J R", "Full-Stack & Biometrics Core", "FastAPI InspireFace AI microservice integration, 512-d feature extraction, real-time WebSocket communication hub, active camera surveillance loop, and alert banners.", c_blue, c_blue_light, c_blue_dark),
        ("Suhas B R", "Security & Steganography", "Client-side Web Crypto API pipeline (AES-256-GCM), friendship secret agreement with PBKDF2 salted hashing, non-exportable IndexedDB keys, and randomized LSB engine.", c_emerald, c_emerald_light, c_emerald),
        ("Dhruva Kumar V M", "Frontend Lead & UI/UX", "Modern React 19 SPA architecture, responsive Tailwind styling with dark/light themes, WebRTC camera lifecycle management, attachment uploading, and master report documentation.", c_purple, c_purple_light, c_purple),
    ]

    col_w = Inches(2.78)
    for i, (m_name, m_role, m_desc, m_accent, m_bg_badge, m_txt_badge) in enumerate(members):
        m_x = Inches(0.95) + i * Inches(2.88)
        m_box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, m_x, Inches(3.55), col_w, Inches(2.8))
        m_box.fill.solid()
        m_box.fill.fore_color.rgb = c_card_subtle
        m_box.line.color.rgb = c_card_border
        m_box.line.width = Pt(1)

        tf_m = m_box.text_frame
        tf_m.word_wrap = True

        p1 = tf_m.paragraphs[0]
        r_name = p1.add_run()
        r_name.text = m_name + "\n"
        r_name.font.bold = True
        r_name.font.size = Pt(11)
        r_name.font.color.rgb = c_text_primary

        r_role = p1.add_run()
        r_role.text = m_role + "\n\n"
        r_role.font.bold = True
        r_role.font.size = Pt(8.5)
        r_role.font.color.rgb = m_accent

        p2 = tf_m.add_paragraph()
        r_desc = p2.add_run()
        r_desc.text = m_desc
        r_desc.font.size = Pt(8)
        r_desc.font.color.rgb = c_text_secondary

    # Interactive Jump Buttons at Bottom of Title Slide
    lbl_j = s1.shapes.add_textbox(Inches(0.8), Inches(6.75), Inches(3.2), Inches(0.4))
    lbl_j.text_frame.paragraphs[0].add_run().text = "Interactive Presentation Navigation:"
    lbl_j.text_frame.paragraphs[0].runs[0].font.bold = True
    lbl_j.text_frame.paragraphs[0].runs[0].font.size = Pt(9.5)
    lbl_j.text_frame.paragraphs[0].runs[0].font.color.rgb = c_text_muted

    add_jump_button(s1, Inches(4.0), Inches(6.72), Inches(1.5), Inches(0.36), "1. Architecture →", s6, bg_color=c_blue_light, border_color=c_blue, text_color=c_blue_dark)
    add_jump_button(s1, Inches(5.6), Inches(6.72), Inches(1.5), Inches(0.36), "2. Features →", s8, bg_color=c_teal_light, border_color=c_teal, text_color=c_teal_dark)
    add_jump_button(s1, Inches(7.2), Inches(6.72), Inches(1.5), Inches(0.36), "3. Testing (20/20) →", s9, bg_color=c_emerald_light, border_color=c_emerald, text_color=c_emerald)
    add_jump_button(s1, Inches(8.8), Inches(6.72), Inches(1.5), Inches(0.36), "4. Screenshots →", s10, bg_color=c_amber_light, border_color=c_amber, text_color=c_amber)
    add_jump_button(s1, Inches(10.4), Inches(6.72), Inches(2.133), Inches(0.36), "⚡ Open Live App ↗", s14, bg_color=c_teal, border_color=c_teal_dark, text_color=c_white)

    # =========================================================================
    # SLIDE 2: EXECUTIVE SUMMARY & PROJECT JOURNEY (Phases 1 → 4)
    # =========================================================================
    add_interactive_header(s2, 1, "Executive Summary: The Complete Journey (Phases 1 → 4)")

    phases = [
        ("Phase 1: Initial Analysis", "IDEATION & CARDBOARD", c_teal_dark, c_teal_light, [
            "Conducted offline whiteboard brainstorming to uncover communication gaps",
            "Conceived a 3-tier defense architecture on a physical cardboard sheet",
            "Evaluated & rejected client WASM vision due to battery drain & inspection risk",
            "Established core ethical privacy policy: Never store raw user facial images",
            "Designed randomized PRNG-based LSB steganography with CVS1 magic tag",
            "Created mathematical threat model for physical over-the-shoulder eavesdropping"
        ]),
        ("Phase 2: Core Engineering", "IMPLEMENTATION SPRINTS", c_blue_dark, c_blue_light, [
            "Sprint 1: Built FastAPI InspireFace microservice with 512-d embeddings",
            "Sprint 2: Built WebSocket instant chat hub & Multer 25MB attachment streaming",
            "Sprint 3: Integrated WebCrypto AES-GCM + non-exportable IndexedDB keys",
            "Sprint 4: Implemented active camera surveillance loop & 2s timeout aborts",
            "Engineered Express 5 single-port reverse proxy serving static Vite SPA",
            "Implemented real-time warning banners for multi-face and bystander detection"
        ]),
        ("Phase 3: Testing & Hardening", "VERIFICATION & QA", c_amber, c_amber_light, [
            "Executed 20 comprehensive test cases across 7 modules (100% Pass Rate)",
            "Validated mandatory single-face enrollment & multi-face rejection (HTTP 422)",
            "Tested active camera surveillance loop: auto-conceals text within 1.5s",
            "Verified boundary limits: 25MB attachment quota & 2,000ms AbortSignal timeouts",
            "Proved database ephemeral self-healing: restores clean original carrier PNG",
            "Validated zero plaintext exposure across network wire, database, and client"
        ]),
        ("Phase 4: Cloud Delivery", "PRODUCTION LAUNCH", c_emerald, c_emerald_light, [
            "Successfully deployed live single-domain container service on Render PaaS",
            "Integrated Supabase PostgreSQL 16 cloud database with SSL connection pooling",
            "Configured automated CI/CD continuous deployment pipeline from GitHub main",
            "Achieved sub-80ms WebSocket delivery latency & sub-70ms facial check",
            "Produced 8-page comprehensive master documentation PDF & interactive deck",
            "Delivered fully functional, battle-tested solution ready for real-world adoption"
        ]),
    ]

    for i, (p_title, p_badge, p_accent, p_bg_badge, p_items) in enumerate(phases):
        card_x = Inches(0.8) + i * Inches(2.98)
        c_box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, card_x, Inches(1.35), Inches(2.82), Inches(5.6))
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = c_card
        c_box.line.color.rgb = p_accent
        c_box.line.width = Pt(1.5)

        tf = c_box.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        r1 = p.add_run()
        r1.text = p_title + "\n"
        r1.font.bold = True
        r1.font.size = Pt(12)
        r1.font.color.rgb = c_text_primary

        r2 = p.add_run()
        r2.text = f"[{p_badge}]\n\n"
        r2.font.bold = True
        r2.font.size = Pt(8)
        r2.font.color.rgb = p_accent

        for item in p_items:
            pi = tf.add_paragraph()
            pi.space_after = Pt(6)
            ri = pi.add_run()
            ri.text = "• " + item
            ri.font.size = Pt(8)
            ri.font.color.rgb = c_text_secondary

    # =========================================================================
    # SLIDE 3: PROBLEM STATEMENT & COMPREHENSIVE THREAT ANALYSIS
    # =========================================================================
    add_interactive_header(s3, 2, "Problem Statement & The 'Last-Inch' Vulnerability Gap")

    # Problem Statement Callout Card
    ps_card = add_card(s3, Inches(0.8), Inches(1.35), Inches(11.733), Inches(1.4), title="ORIGINAL PROBLEM STATEMENT ASSIGNED TO TEAM 65", title_color=c_teal_dark, bg_color=c_white)
    tf_ps = ps_card.text_frame
    p_ps = tf_ps.add_paragraph()
    p_ps.space_before = Pt(16)
    r_ps = p_ps.add_run()
    r_ps.text = '"Design and engineer a real-time secure communication system that eliminates vulnerabilities in both data transmission and physical message presentation. The application must guarantee that confidential exchanges remain fully protected from over-the-shoulder eavesdropping, unauthorized physical device access, and persistent forensic extraction, while providing zero-knowledge verifiable privacy and high usability."'
    r_ps.font.italic = True
    r_ps.font.size = Pt(10)
    r_ps.font.color.rgb = c_text_secondary

    # 4 Vulnerability Dimension Cards
    gaps = [
        ("1. The 'Last-Inch' Presentation Vulnerability", "Conventional E2EE (Signal, WhatsApp) encrypts data in flight via TLS and Signal Protocol. However, the instant a message decrypts onto the recipient's screen, it exists in plain sight. In public environments—coffee shops, airports, shared offices, and public transit—shoulder surfers and bystanders compromise confidentiality completely.", c_amber, "CRITICAL PRESENTATION FLAW"),
        ("2. Static vs Continuous Active Authentication", "Modern operating systems and messaging apps authenticate biometrics (Face ID, Touch ID, PIN) only ONCE when launching the application. If an unlocked smartphone or laptop is left unattended on a desk, passed to a coworker, or snatched, any unauthorized person has unrestricted access to sensitive chats.", c_amber, "IDENTITY SESSION HIJACKING"),
        ("3. Persistent Forensic Media Residue & Caches", "Standard messaging apps write received attachments to local device flash storage, OS media caches, and SQLite databases. Even after 'deleting' messages or enabling disappearing timers, forensic extraction tools routinely recover sensitive image attachments and thumbnail artifacts from raw storage blocks.", c_amber, "FORENSIC DATA LEAKAGE"),
        ("4. Traffic Observability & Lack of Deniability", "Transmitting raw encrypted binary blocks openly alerts network firewalls, internet service providers (ISPs), and eavesdroppers that sensitive, confidential communication is taking place. Users completely lack plausible deniability without an innocuous, covert carrier mechanism.", c_amber, "TRAFFIC METADATA EXPOSURE")
    ]

    for i, (g_title, g_desc, g_accent, g_badge) in enumerate(gaps):
        gx = Inches(0.8) + (i % 2) * Inches(5.95)
        gy = Inches(2.95) + (i // 2) * Inches(2.1)
        add_card(s3, gx, gy, Inches(5.78), Inches(1.95), title=g_title, title_color=c_text_primary, badge_text=g_badge, badge_bg=c_amber_light, badge_color=c_amber)

        tb = s3.shapes.add_textbox(gx + Inches(0.18), gy + Inches(0.48), Inches(5.42), Inches(1.35))
        tb.text_frame.word_wrap = True
        p = tb.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = g_desc
        r.font.size = Pt(8.5)
        r.font.color.rgb = c_text_secondary

    # =========================================================================
    # SLIDE 4: PROPOSED SOLUTION: ZERO-TRUST DEFENSE-IN-DEPTH
    # =========================================================================
    add_interactive_header(s4, 3, "Proposed Solution: Multi-Layered Zero-Trust Architecture")

    pillars = [
        ("AI Biometric Active Shield", "Continuous Vision Surveillance", c_teal_dark, c_teal_light, [
            "InspireFace Deep Learning Computer Vision SDK",
            "512-dimensional normalized facial embedding vectors",
            "Real-time Cosine Similarity verification (>= 0.70 threshold)",
            "Enforces exactly 1 human face present in frame at all times",
            "Continuous 1.5s background surveillance loop during viewing",
            "Auto-conceals message text instantly if viewer turns away",
            "Displays warning banner if secondary shoulder-surfer appears"
        ]),
        ("Randomized Steganography", "Plausible Deniability Carrier", c_blue_dark, c_blue_light, [
            "Disguises encrypted payloads inside ordinary lossless PNG images",
            "24-byte custom header: CVS1 magic tag + 16-byte seed + 4-byte size",
            "16-byte cryptographically secure random PRNG seed token",
            "Uniform pseudo-random pixel channel scattering across RGB",
            "Zero perceptible visual distortion in carrier image histograms",
            "Defeats automated statistical and visual steganalysis tools",
            "Provides plausible deniability against network payload sniffers"
        ]),
        ("Client-Side WebCrypto", "Zero-Knowledge Key Agreement", c_emerald, c_emerald_light, [
            "Out-of-band mutual friendship passphrase agreed by peers",
            "Client WebCrypto API PBKDF2-SHA256 (100,000 iterations)",
            "AES-GCM-256 authenticated encryption with unique 12-byte IVs",
            "Derived keys stored as non-exportable CryptoKeys in IndexedDB",
            "Server receives only a salted verifier hash for verification",
            "Server never holds plaintext, passphrases, or encryption keys",
            "Protects communications even under total server compromise"
        ]),
        ("Forensic Self-Healing", "Ephemeral Database Destruction", c_purple, c_purple_light, [
            "One-time secure message decryption and viewing model",
            "Atomic PostgreSQL database transaction post-consumption",
            "Restores pristine original carrier PNG into attachments table",
            "Permanently nullifies and deletes stego recovery copies",
            "Leaves zero forensic traces of steganography on disk or DB",
            "Guarantees forward secrecy: destroyed messages cannot be recovered",
            "Combines physical viewing security with digital self-healing"
        ]),
    ]

    for i, (p_title, p_sub, p_accent, p_bg_badge, p_items) in enumerate(pillars):
        px = Inches(0.8) + i * Inches(2.98)
        p_box = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px, Inches(1.35), Inches(2.82), Inches(5.6))
        p_box.fill.solid()
        p_box.fill.fore_color.rgb = c_card
        p_box.line.color.rgb = p_accent
        p_box.line.width = Pt(1.5)

        tf = p_box.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        r1 = p.add_run()
        r1.text = p_title + "\n"
        r1.font.bold = True
        r1.font.size = Pt(11.5)
        r1.font.color.rgb = c_text_primary

        r2 = p.add_run()
        r2.text = f"[{p_sub}]\n\n"
        r2.font.bold = True
        r2.font.size = Pt(7.5)
        r2.font.color.rgb = p_accent

        for item in p_items:
            pi = tf.add_paragraph()
            pi.space_after = Pt(5)
            ri = pi.add_run()
            ri.text = "• " + item
            ri.font.size = Pt(7.8)
            ri.font.color.rgb = c_text_secondary

    # =========================================================================
    # SLIDE 5: SYSTEM METHODOLOGY: 8-PHASE WORKFLOW
    # =========================================================================
    add_interactive_header(s5, 4, "Methodology: End-to-End System Workflow (8 Phases)")

    steps = [
        ("01", "Registration & Biometric Consent", "User registers with email, password (bcrypt hashed), and camera capture. FastAPI executes InspireFace enrollment; exactly 1 face required. Generates 512-d template, encrypts via AES-256-GCM, stores in PostgreSQL face_biometrics. Raw photo is immediately purged.", c_teal_dark),
        ("02", "Privacy Friendship Handshake", "Users search for peers via unique 32-character User ID codes (user_id_code). Friend requests are exchanged and accepted. Once accepted, friends agree out-of-band on a shared passphrase, verified via salted PBKDF2 hash on the server.", c_teal_dark),
        ("03", "Client WebCrypto Key Derivation", "Client browser executes Web Crypto API subtle.deriveKey() using PBKDF2-SHA256 (100,000 iterations). Non-exportable AES-GCM 256-bit CryptoKey is stored securely in IndexedDB on the device. Plaintext passphrase is never saved.", c_teal_dark),
        ("04", "Steganographic Carrier Encoding", "Sender toggles 'Private image message', selects PNG and enters text. Text is encrypted with AES-GCM (12-byte IV). Payload is embedded into PNG via randomized LSB (stego/embed). Stego PNG and recovery original sent to server.", c_teal_dark),
        ("05", "Real-Time WebSocket Delivery", "Node.js Express backend stores attachment BYTEA in PostgreSQL and broadcasts payload over WebSocket (ws://). Recipient UI updates instantaneously with the message card displaying the carrier image within 80ms.", c_blue_dark),
        ("06", "Biometric Recipient Gating", "Recipient clicks decrypt icon. WebRTC activates camera, transmits frame to POST /face/verify. AI validates cosine similarity >= 0.70 and ensures face count == 1. Payload is extracted, decrypted via IndexedDB key, and revealed.", c_blue_dark),
        ("07", "Active Surveillance Loop & Timeout", "Continuous verification loop samples video frames every 1,500ms with a 2,000ms AbortSignal timeout. If face count != 1 or cosine similarity falls below threshold, text is concealed immediately with an on-screen alert banner.", c_blue_dark),
        ("08", "Ephemeral Forensic Self-Healing", "Database marks stego message consumed. PostgreSQL updates attachments table with the original clean PNG and nullifies original_png, permanently removing stego bytes and eliminating post-read forensic discovery.", c_blue_dark),
    ]

    for i, (num, s_title, s_desc, s_color) in enumerate(steps):
        col = i % 4
        row = i // 4
        sx = Inches(0.8) + col * Inches(2.98)
        sy = Inches(1.35) + row * Inches(2.7)

        s_box = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx, sy, Inches(2.82), Inches(2.55))
        s_box.fill.solid()
        s_box.fill.fore_color.rgb = c_card
        s_box.line.color.rgb = s_color
        s_box.line.width = Pt(1)

        tf = s_box.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        r_num = p.add_run()
        r_num.text = f"PHASE {num}\n"
        r_num.font.bold = True
        r_num.font.size = Pt(8.5)
        r_num.font.color.rgb = s_color

        r_title = p.add_run()
        r_title.text = s_title + "\n\n"
        r_title.font.bold = True
        r_title.font.size = Pt(10)
        r_title.font.color.rgb = c_text_primary

        p_desc = tf.add_paragraph()
        r_desc = p_desc.add_run()
        r_desc.text = s_desc
        r_desc.font.size = Pt(7.8)
        r_desc.font.color.rgb = c_text_secondary

    # =========================================================================
    # SLIDE 6: SYSTEM ARCHITECTURE (Decoupled 4-Tier Cloud Design)
    # =========================================================================
    add_interactive_header(s6, 5, "System Architecture: Decoupled 4-Tier Cloud Architecture")

    # Tier 1: Client
    t1 = add_card(s6, Inches(0.8), Inches(1.35), Inches(11.733), Inches(1.35), title="TIER 1: PRESENTATION LAYER (React 19 + TypeScript + Vite + Tailwind CSS)", title_color=c_teal_dark, border_color=c_teal)
    tb1 = s6.shapes.add_textbox(Inches(1.0), Inches(1.75), Inches(11.333), Inches(0.85))
    tb1.text_frame.word_wrap = True
    p1 = tb1.text_frame.paragraphs[0]
    p1.add_run().text = "• WebCrypto API: Client AES-GCM-256 derivation & non-exportable IndexedDB key persistence | PBKDF2-SHA256 (100k rounds)\n• WebRTC MediaStream: Hardware webcam feed capture & active 1.5s background surveillance loop | AbortSignal 2s timeout protection\n• WebSocket Client (ws://): Instant bidirectional message synchronization, typing broadcasts, online presence indicators & error banners"
    p1.runs[0].font.size = Pt(8.5)
    p1.runs[0].font.color.rgb = c_text_secondary

    # Connector arrow 1
    arrow1 = s6.shapes.add_textbox(Inches(4.5), Inches(2.72), Inches(4.333), Inches(0.32))
    p_a1 = arrow1.text_frame.paragraphs[0]
    p_a1.alignment = PP_ALIGN.CENTER
    r_a1 = p_a1.add_run()
    r_a1.text = "▼ RESTful HTTP JSON APIs & Bi-Directional WebSockets (Port 443/5000) ▼"
    r_a1.font.bold = True
    r_a1.font.size = Pt(8)
    r_a1.font.color.rgb = c_teal_dark

    # Tier 2: Gateway
    t2 = add_card(s6, Inches(0.8), Inches(3.05), Inches(11.733), Inches(1.35), title="TIER 2: GATEWAY & APPLICATION SERVER (Node.js + Express 5)", title_color=c_blue_dark, border_color=c_blue)
    tb2 = s6.shapes.add_textbox(Inches(1.0), Inches(3.45), Inches(11.333), Inches(0.85))
    tb2.text_frame.word_wrap = True
    p2 = tb2.text_frame.paragraphs[0]
    p2.add_run().text = "• Gateway & Reverse Proxy: Serves Vite compiled SPA dist/ statically and handles catch-all SPA routing (/*splat)\n• Security Middleware: JWT bearer authentication, bcrypt password verification (10 rounds), AES-256-GCM face template encryptor\n• Data Handlers: Multer 25MB attachment streaming & node-postgres (pg) connection pool with dynamic SSL for Supabase"
    p2.runs[0].font.size = Pt(8.5)
    p2.runs[0].font.color.rgb = c_text_secondary

    # Connector arrow 2 (Split)
    arrow2a = s6.shapes.add_textbox(Inches(1.5), Inches(4.42), Inches(4.0), Inches(0.32))
    p_a2a = arrow2a.text_frame.paragraphs[0]
    p_a2a.alignment = PP_ALIGN.CENTER
    r_a2a = p_a2a.add_run()
    r_a2a.text = "▼ Internal HTTP Proxy (:8001) ▼"
    r_a2a.font.bold = True
    r_a2a.font.size = Pt(8)
    r_a2a.font.color.rgb = c_text_muted

    arrow2b = s6.shapes.add_textbox(Inches(7.5), Inches(4.42), Inches(4.0), Inches(0.32))
    p_a2b = arrow2b.text_frame.paragraphs[0]
    p_a2b.alignment = PP_ALIGN.CENTER
    r_a2b = p_a2b.add_run()
    r_a2b.text = "▼ Encrypted SSL TCP Pooler (:5432) ▼"
    r_a2b.font.bold = True
    r_a2b.font.size = Pt(8)
    r_a2b.font.color.rgb = c_text_muted

    # Tier 3: AI Microservice
    t3 = add_card(s6, Inches(0.8), Inches(4.75), Inches(5.75), Inches(2.2), title="TIER 3: AI & STEGO ENGINE (Python FastAPI)", title_color=c_emerald, border_color=c_emerald)
    tb3 = s6.shapes.add_textbox(Inches(1.0), Inches(5.15), Inches(5.35), Inches(1.7))
    tb3.text_frame.word_wrap = True
    p3 = tb3.text_frame.paragraphs[0]
    p3.add_run().text = "• InspireFace SDK ('Pikachu' Model): Deep face detection & alignment\n• 512-d Embedding Extraction: High-dimension normalized vector creation\n• Cosine Similarity Verification: Match score threshold >= 0.70\n• Session Locking (_session_lock): Thread-safe C-bindings execution\n• Randomized LSB Engine: NumPy PRNG seeded bit dispersal across RGB"
    p3.runs[0].font.size = Pt(8)
    p3.runs[0].font.color.rgb = c_text_secondary

    # Tier 4: Database
    t4 = add_card(s6, Inches(6.78), Inches(4.75), Inches(5.75), Inches(2.2), title="TIER 4: PERSISTENCE (PostgreSQL 16 Supabase)", title_color=c_purple, border_color=c_purple)
    tb4 = s6.shapes.add_textbox(Inches(6.98), Inches(5.15), Inches(5.35), Inches(1.7))
    tb4.text_frame.word_wrap = True
    p4 = tb4.text_frame.paragraphs[0]
    p4.add_run().text = "• users & face_biometrics: Encrypted template ciphertexts + nonces\n• friend_requests & friendship_secrets: Salted PBKDF2 hash verifiers\n• messages & attachments: BYTEA binary blocks up to 25MB\n• stego_messages: Ephemeral carrier recovery copies with CASCADE\n• Ephemeral Cleanup Triggers: Atomic transactions for self-healing"
    p4.runs[0].font.size = Pt(8)
    p4.runs[0].font.color.rgb = c_text_secondary

    # =========================================================================
    # SLIDE 7: TECHNOLOGIES, FRAMEWORKS & TOOLING STACK
    # =========================================================================
    add_interactive_header(s7, 6, "Technologies, Frameworks & Tooling Stack")

    tech_domains = [
        ("Frontend UI & Client Security", c_teal_dark, c_teal_light, [
            ("React 19 & TypeScript", "Component-based reactive SPA, strict type definitions, fast re-renders"),
            ("Vite 8 Bundler", "Modern ES module frontend tooling, optimized production minification"),
            ("Tailwind CSS", "Ergonomic responsive styling, dark/light theme support, accessible layout"),
            ("Web Crypto API", "Native browser PBKDF2-SHA256 key derivation & AES-GCM-256 ciphers"),
            ("IndexedDB API", "Browser-isolated, non-exportable cryptographic key storage"),
            ("HTML5 WebRTC", "Real-time webcam MediaStream acquisition & hardware video frame sampling"),
            ("Lucide React Icons", "Clean visual iconography for intuitive navigation and status indicators")
        ]),
        ("Backend & Real-Time Gateway", c_blue_dark, c_blue_light, [
            ("Node.js & Express 5", "Unified RESTful API server, middleware pipeline & static SPA reverse proxy"),
            ("WebSockets (`ws`)", "Bi-directional real-time messaging, typing indicators & presence state"),
            ("JSON Web Tokens (JWT)", "Stateless, digitally signed bearer session token authentication"),
            ("bcrypt (10 rounds)", "Cryptographic salted password hashing resisting brute-force attacks"),
            ("Multer Middleware", "Memory-buffered multipart attachment handling supporting files up to 25MB"),
            ("node-postgres (`pg`)", "High-performance connection pool with dynamic SSL certificate handling")
        ]),
        ("AI Vision & Steganography", c_emerald, c_emerald_light, [
            ("Python 3.12 & FastAPI", "Asynchronous, high-throughput microservice for AI computation"),
            ("InspireFace SDK", "State-of-the-art deep facial recognition & 512-d feature extraction"),
            ("OpenCV (`cv2`)", "Raw frame matrix decoding, color space transformation & alignment"),
            ("NumPy 2.4", "Vectorized array manipulation & PRNG bit scattering across RGB channels"),
            ("Pillow (`PIL`)", "Lossless RGB PNG image encoding, decoding & dimension verification"),
            ("Python Secrets", "Cryptographically secure 16-byte PRNG token generation for stego seeds")
        ]),
        ("Database, Cloud & DevOps", c_purple, c_purple_light, [
            ("PostgreSQL 16", "ACID-compliant relational database with binary BYTEA storage"),
            ("Supabase Managed DB", "Cloud-hosted managed database with transaction pooler support"),
            ("Render PaaS", "Containerized single-domain cloud hosting with automated builds"),
            ("Git & GitHub", "Distributed version control, branch isolation, and deployment hooks"),
            ("ReportLab & PyMuPDF", "Automated publication-grade PDF generation with dynamic page numbering"),
            ("python-pptx", "Automated widescreen interactive presentation generation")
        ]),
    ]

    for i, (dom_name, dom_accent, dom_bg_badge, dom_items) in enumerate(tech_domains):
        dx = Inches(0.8) + i * Inches(2.98)
        d_box = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, dx, Inches(1.35), Inches(2.82), Inches(5.6))
        d_box.fill.solid()
        d_box.fill.fore_color.rgb = c_card
        d_box.line.color.rgb = dom_accent
        d_box.line.width = Pt(1.5)

        tf = d_box.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = dom_name + "\n\n"
        r.font.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = c_text_primary

        for tech, desc in dom_items:
            pt = tf.add_paragraph()
            pt.space_after = Pt(5)
            rt = pt.add_run()
            rt.text = tech + "\n"
            rt.font.bold = True
            rt.font.size = Pt(8)
            rt.font.color.rgb = dom_accent

            rd = pt.add_run()
            rd.text = desc
            rd.font.size = Pt(7.2)
            rd.font.color.rgb = c_text_secondary

    # =========================================================================
    # SLIDE 8: PHASE 2 IMPLEMENTATION: 5 CORE FEATURES
    # =========================================================================
    add_interactive_header(s8, 7, "Phase 2 Implementation: 5 Core Engineering Features")

    innovations = [
        ("Feature 1: Zero-Knowledge Biometric Enrollment", "Mandatory registration captures exactly 1 face using InspireFace. Extracts 512-d normalized floating point embeddings. The Node.js backend encrypts the vector with AES-256-GCM using FACE_ENCRYPTION_KEY before database insertion. The raw user photo is purged from memory immediately—never stored on disk or logged.", c_teal_dark, "BIOMETRIC PRIVACY"),
        ("Feature 2: Randomized LSB Steganography Engine", "Disguises encrypted secret text into lossless PNG carriers. Writes a 24-byte header (CVS1 magic tag, 16-byte random PRNG seed, 4-byte payload size) into initial pixels, then scatters encrypted bits randomly across RGB channels via NumPy default_rng. Resulting image exhibits zero perceptible visual degradation or histogram anomalies.", c_blue_dark, "COVERT TRANSPORT"),
        ("Feature 3: Active Face Guardian Surveillance Loop", "Decryption activates recipient webcam. The AI verifies identity against enrolled template (cosine >= 0.70) and guarantees exactly 1 face. A continuous surveillance loop checks the viewer every 1.5s with a 2s AbortSignal timeout. If the user turns away or a second face appears, text is auto-concealed immediately.", c_emerald, "PHYSICAL SHIELD"),
        ("Feature 4: Single-Port Cloud Bundling on Render", "Cloud PaaS provides only one public port ($PORT). Team Veltrox engineered a root build script compiling Vite SPA and Express server into a unified artifact. Express 5 acts as reverse proxy, serving React statically from dist/ and handling catch-all SPA routing seamlessly while hosting WebSockets.", c_amber, "CLOUD DEVOPS"),
        ("Feature 5: Forensic Database Self-Healing", "Traditional apps leave persistent traces of decrypted media. Chatvia implements atomic self-healing: upon single-view decryption, PostgreSQL updates attachments with the original clean PNG and nullifies original_png, removing any forensic trace of steganography from disk and database backups.", c_purple, "EPHEMERAL HEALING"),
    ]

    for i, (f_title, f_desc, f_accent, f_badge) in enumerate(innovations):
        fy = Inches(1.35) + i * Inches(1.12)
        f_card = add_card(s8, Inches(0.8), fy, Inches(11.733), Inches(1.02), title=f_title, title_color=c_text_primary, badge_text=f_badge, badge_bg=c_card_subtle, badge_color=f_accent)

        tb = s8.shapes.add_textbox(Inches(1.0), fy + Inches(0.38), Inches(11.333), Inches(0.58))
        tb.text_frame.word_wrap = True
        p = tb.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = f_desc
        r.font.size = Pt(8.3)
        r.font.color.rgb = c_text_secondary

    # =========================================================================
    # SLIDE 9: PHASE 3 TESTING & VERIFICATION (100% PASS RATE)
    # =========================================================================
    add_interactive_header(s9, 8, "Phase 3 Testing: 20 Test Cases Executed (100% Pass Rate)")

    # Metrics Bar
    metrics = [
        ("Total Test Cases", "20", c_blue_dark, c_blue_light),
        ("Execution Pass", "100%", c_emerald, c_emerald_light),
        ("Face Verify Latency", "< 70 ms", c_teal_dark, c_teal_light),
        ("WebSocket Latency", "< 80 ms", c_teal_dark, c_teal_light),
        ("Max File Upload", "25 MB", c_amber, c_amber_light),
        ("Network Abort Timeout", "2,000 ms", c_purple, c_purple_light),
    ]

    for i, (m_lbl, m_val, m_col, m_bg) in enumerate(metrics):
        mx = Inches(0.8) + i * Inches(1.98)
        m_box = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, mx, Inches(1.35), Inches(1.82), Inches(0.85))
        m_box.fill.solid()
        m_box.fill.fore_color.rgb = m_bg
        m_box.line.color.rgb = m_col
        m_box.line.width = Pt(1)

        tf = m_box.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p1 = tf.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        r1 = p1.add_run()
        r1.text = m_val + "\n"
        r1.font.bold = True
        r1.font.size = Pt(15)
        r1.font.color.rgb = m_col

        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = m_lbl
        r2.font.bold = True
        r2.font.size = Pt(7.5)
        r2.font.color.rgb = c_text_secondary

    # Test Cases Table
    t_shape = s9.shapes.add_table(7, 5, Inches(0.8), Inches(2.35), Inches(11.733), Inches(4.6))
    table = t_shape.table
    table.columns[0].width = Inches(1.1)
    table.columns[1].width = Inches(1.8)
    table.columns[2].width = Inches(3.4)
    table.columns[3].width = Inches(4.333)
    table.columns[4].width = Inches(1.1)

    headers = ["Test ID", "Feature Module", "Condition / Input Tested", "Expected & Verified Outcome", "Status"]
    for col_idx, h_text in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = c_teal_dark
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = h_text
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = c_white

    sample_tests = [
        ("TC-REG-02", "Biometric Enrollment", "User attempts registration with 2 faces visible in webcam frame", "HTTP 422 rejected; error 'Capture exactly one face, clearly visible'; blocked", "PASS (100%)"),
        ("TC-KEY-01", "Friendship Key Derivation", "Both friends input matching shared passphrase out-of-band", "PBKDF2 salted hash verified; non-exportable AES-GCM key cached in IndexedDB", "PASS (100%)"),
        ("TC-STEG-01", "Randomized LSB Stego", "500-char private text embedded into 1.2MB lossless PNG carrier", "Header CVS1 written; payload scattered via PRNG; zero visual distortion", "PASS (100%)"),
        ("TC-BIO-03", "Shoulder Surfer Detection", "Second person enters webcam view while recipient is viewing text", "Active surveillance detects face_count == 2; text auto-concealed immediately", "PASS (100%)"),
        ("TC-BIO-05", "Timeout Protection", "Simulated network latency > 2,000ms during face verification call", "AbortSignal.timeout(2000) cancels request; text concealed safely with banner", "PASS (100%)"),
        ("TC-HEAL-01", "Forensic Self-Healing", "Message decrypted and successfully viewed once by authorized recipient", "Database restores original clean PNG to attachments; recovery copy deleted", "PASS (100%)"),
    ]

    for row_idx, row_data in enumerate(sample_tests, start=1):
        for col_idx, text in enumerate(row_data):
            cell = table.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = c_white if row_idx % 2 == 1 else c_card_subtle
            p = cell.text_frame.paragraphs[0]
            if col_idx in (0, 4):
                p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = text
            r.font.size = Pt(8)
            if col_idx == 4:
                r.font.bold = True
                r.font.color.rgb = c_emerald
            elif col_idx == 0:
                r.font.bold = True
                r.font.color.rgb = c_text_primary
            else:
                r.font.color.rgb = c_text_secondary

    # =========================================================================
    # SLIDE 10: APPLICATION SCREENSHOTS - DESKTOP DASHBOARD
    # =========================================================================
    add_interactive_header(s10, 9, "Application Visual Proof: Desktop Chatvia Dashboard")

    img1_path = "Screenshots/Screenshot 2026-10-29 034058.png"
    if os.path.exists(img1_path):
        add_card(s10, Inches(0.8), Inches(1.35), Inches(8.3), Inches(5.6), border_color=c_teal)
        s10.shapes.add_picture(img1_path, Inches(0.9), Inches(1.45), Inches(8.1), Inches(5.4))

        # Annotations Card
        add_card(s10, Inches(9.3), Inches(1.35), Inches(3.233), Inches(5.6), title="DESKTOP FEATURES", title_color=c_text_primary)
        tb = s10.shapes.add_textbox(Inches(9.45), Inches(1.8), Inches(2.933), Inches(5.0))
        tb.text_frame.word_wrap = True

        callouts = [
            ("Active Contact Directory", "Bilateral friend relationships using unique 32-char User ID codes. Displays online status in real-time."),
            ("Real-Time Message Stream", "Instant WebSocket delivery with sub-80ms broadcast latency. Supports text, files, and private images."),
            ("Steganographic Message Card", "Carriers render as ordinary innocent PNGs. Decrypt trigger requires biometric gating."),
            ("Responsive Dark UI", "Engineered with Tailwind CSS and Lucide icons for maximum ergonomics and focus."),
            ("Live Production URL", "Deployed and tested live on Render PaaS at chatvia-app-k7o7.onrender.com")
        ]
        for c_title, c_text in callouts:
            p_c = tb.text_frame.add_paragraph()
            p_c.space_after = Pt(8)
            r1 = p_c.add_run()
            r1.text = "• " + c_title + ":\n"
            r1.font.bold = True
            r1.font.size = Pt(8.5)
            r1.font.color.rgb = c_teal_dark

            r2 = p_c.add_run()
            r2.text = c_text
            r2.font.size = Pt(7.5)
            r2.font.color.rgb = c_text_secondary

    # =========================================================================
    # SLIDE 11: APPLICATION SCREENSHOTS - MOBILE & BIOMETRICS
    # =========================================================================
    add_interactive_header(s11, 10, "Application Visual Proof: Biometric Enrollment & Stego Sender")

    img2_path = "Screenshots/Screenshot 2026-10-29 034122.png"
    img3_path = "Screenshots/Screenshot 2026-10-29 034136.png"

    if os.path.exists(img2_path) and os.path.exists(img3_path):
        add_card(s11, Inches(0.8), Inches(1.35), Inches(3.4), Inches(5.6), title="BIOMETRIC ENROLLMENT", title_color=c_teal_dark, border_color=c_teal)
        s11.shapes.add_picture(img2_path, Inches(0.95), Inches(1.8), Inches(3.1), Inches(5.0))

        add_card(s11, Inches(4.4), Inches(1.35), Inches(3.4), Inches(5.6), title="PRIVATE STEGO MESSAGE", title_color=c_blue_dark, border_color=c_blue)
        s11.shapes.add_picture(img3_path, Inches(4.55), Inches(1.8), Inches(3.1), Inches(5.0))

        add_card(s11, Inches(8.0), Inches(1.35), Inches(4.533), Inches(5.6), title="MOBILE & SECURITY WALKTHROUGH", title_color=c_text_primary)
        tb = s11.shapes.add_textbox(Inches(8.2), Inches(1.8), Inches(4.133), Inches(5.0))
        tb.text_frame.word_wrap = True

        walkthrough = [
            ("Mandatory Biometric Registration (Figure 2)", "Users cannot sign up without explicit camera consent and single-face capture. InspireFace validates alignment, eliminates multi-face fraud, and outputs a 512-d embedding. The raw photograph is instantly destroyed."),
            ("Master Key Encryption Before Persistence", "Biometric templates are encrypted with AES-256-GCM via the backend's secret master key prior to PostgreSQL storage, preventing data leakage in database dumps."),
            ("Private Image Message Composition (Figure 3)", "Users toggle 'Private image message' in the chat composer. An innocent carrier PNG is selected, and secret text is typed. The client encrypts the text with WebCrypto before sending to the stego engine."),
            ("Friendship Key Management & Rotation", "Friends agree on a shared passphrase once. Derived keys persist in IndexedDB without ever being sent across the wire. Key rotation is blocked while any private message remains unopened.")
        ]
        for w_title, w_desc in walkthrough:
            p_w = tb.text_frame.add_paragraph()
            p_w.space_after = Pt(10)
            r1 = p_w.add_run()
            r1.text = "• " + w_title + ":\n"
            r1.font.bold = True
            r1.font.size = Pt(8.5)
            r1.font.color.rgb = c_teal_dark

            r2 = p_w.add_run()
            r2.text = w_desc
            r2.font.size = Pt(7.8)
            r2.font.color.rgb = c_text_secondary

    # =========================================================================
    # SLIDE 12: CHALLENGES FACED & SOLUTIONS
    # =========================================================================
    add_interactive_header(s12, 11, "Challenges Faced & Engineering Solutions")

    challenges = [
        ("1. Render Single-Port Limitation ($PORT)", "Cloud PaaS provides only one public web port ($PORT), but the architecture requires a React SPA, Express REST API, and WebSocket channels.", "Engineered a unified root build script (`npm run build`) compiling both client and server. Express 5 acts as reverse proxy, serving Vite's compiled dist/ directory statically and routing API/WS paths cleanly.", c_teal_dark, "SOLVED: REVERSE PROXY"),
        ("2. Supabase SSL Pooler Rejection", "The Node.js pg driver crashed with SELF_SIGNED_CERT_IN_CHAIN when connecting to the Supabase transaction pooler.", "Implemented dynamic SSL configuration in server/src/db.ts detecting DATABASE_URL and enforcing { rejectUnauthorized: false } for seamless cloud pooler connectivity.", c_blue_dark, "SOLVED: DYNAMIC SSL"),
        ("3. Python 3.12 AI C-Library Wheels", "InspireFace native C-bindings lack prebuilt binaries for Python 3.13, causing segmentation faults on newer runtimes.", "Configured dedicated Python 3.12 environment and introduced a thread synchronization lock (_session_lock) around inference calls to guarantee thread safety.", c_emerald, "SOLVED: THREAD LOCKING"),
        ("4. Camera Stream Lifecycle Leaks", "Background tabs or switching views left hardware webcam tracks active, wasting battery and triggering security indicators.", "Implemented HTML5 visibilitychange listeners and ended event handlers that instantly terminate all media stream tracks when the user changes tabs or closes the card.", c_amber, "SOLVED: TRACK CLEANUP"),
        ("5. Steganographic Visual Artifacts", "Traditional sequential LSB insertion produces noticeable rectangular bands or color shifts in image gradients.", "Developed a randomized NumPy PRNG bit distribution seeded by a 16-byte cryptographic token, scattering bits across RGB channels with zero perceptible change.", c_purple, "SOLVED: PRNG DISPERSAL"),
    ]

    for i, (c_name, c_prob, c_sol, c_accent, c_badge) in enumerate(challenges):
        cy = Inches(1.35) + i * Inches(1.12)
        add_card(s12, Inches(0.8), cy, Inches(11.733), Inches(1.02), title=c_name, title_color=c_text_primary, badge_text=c_badge, badge_bg=c_card_subtle, badge_color=c_accent)

        tb = s12.shapes.add_textbox(Inches(1.0), cy + Inches(0.35), Inches(11.333), Inches(0.6))
        tb.text_frame.word_wrap = True
        p = tb.text_frame.paragraphs[0]

        r1 = p.add_run()
        r1.text = "Challenge: "
        r1.font.bold = True
        r1.font.size = Pt(8)
        r1.font.color.rgb = c_amber

        r2 = p.add_run()
        r2.text = c_prob + "  "
        r2.font.size = Pt(8)
        r2.font.color.rgb = c_text_secondary

        r3 = p.add_run()
        r3.text = "→ Solution: "
        r3.font.bold = True
        r3.font.size = Pt(8)
        r3.font.color.rgb = c_emerald

        r4 = p.add_run()
        r4.text = c_sol
        r4.font.size = Pt(8)
        r4.font.color.rgb = c_text_primary

    # =========================================================================
    # SLIDE 13: RESULTS, LIMITATIONS & FUTURE ROADMAP
    # =========================================================================
    add_interactive_header(s13, 12, "Results, Outcomes, Limitations & Future Roadmap")

    # Left Card: Results & Outcomes
    add_card(s13, Inches(0.8), Inches(1.35), Inches(5.75), Inches(5.6), title="ACHIEVEMENTS & MEASURED RESULTS", title_color=c_emerald, border_color=c_emerald)
    tb_res = s13.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(5.35), Inches(5.0))
    tb_res.text_frame.word_wrap = True

    results = [
        ("Sub-80ms Real-Time Delivery", "WebSocket broadcasts achieve instantaneous message synchronization between clients under real-world network loads."),
        ("< 70ms Biometric Inference", "InspireFace 512-d feature extraction and cosine verification execute in under 70ms on standard CPUs."),
        ("Zero-Knowledge Privacy", "Zero raw facial images stored; zero plaintext stored in database; encryption keys never touch the server."),
        ("100% Comprehensive Test Pass", "All 20 test cases validated across registration, auth, keys, steganography, biometrics, and database self-healing."),
        ("Production Cloud PaaS Deployment", "Publicly hosted and accessible 24/7 on Render at chatvia-app-k7o7.onrender.com")
    ]
    for r_title, r_desc in results:
        p = tb_res.text_frame.add_paragraph()
        p.space_after = Pt(8)
        r1 = p.add_run()
        r1.text = "✔ " + r_title + ":\n"
        r1.font.bold = True
        r1.font.size = Pt(8.8)
        r1.font.color.rgb = c_emerald

        r2 = p.add_run()
        r2.text = r_desc
        r2.font.size = Pt(7.8)
        r2.font.color.rgb = c_text_secondary

    # Right Card: Limitations & Future Roadmap
    add_card(s13, Inches(6.78), Inches(1.35), Inches(5.75), Inches(5.6), title="LIMITATIONS & FUTURE ROADMAP", title_color=c_blue_dark, border_color=c_blue)
    tb_fut = s13.shapes.add_textbox(Inches(6.98), Inches(1.8), Inches(5.35), Inches(5.0))
    tb_fut.text_frame.word_wrap = True

    roadmap = [
        ("Current Limitations", "Requires lossless PNG carrier (JPEG/WebP corrupts LSB bits); 2D camera verification does not detect 3D infrared depth; per-device IndexedDB key requires unlocking once on new browsers."),
        ("Active Liveness & Micro-Movement", "Future enhancement to track blink frequency, head-pose jitter, and micro-movements to defeat high-res photograph spoofing attacks."),
        ("Zero-Knowledge Biometric Proofs (zk-SNARKs)", "Enabling clients to mathematically prove biometric ownership without transmitting encrypted embedding vectors to the server."),
        ("Audio & Video Steganography", "Expanding carrier support to lossless FLAC audio and WebM video streams for rich confidential multimedia data transfers."),
        ("Decentralized P2P WebRTC Data Mesh", "Direct device-to-device encrypted data pipes bypassing centralized servers for extreme high-confidentiality scenarios.")
    ]
    for m_title, m_desc in roadmap:
        p = tb_fut.text_frame.add_paragraph()
        p.space_after = Pt(8)
        r1 = p.add_run()
        r1.text = "✦ " + m_title + ":\n"
        r1.font.bold = True
        r1.font.size = Pt(8.8)
        r1.font.color.rgb = c_blue_dark

        r2 = p.add_run()
        r2.text = m_desc
        r2.font.size = Pt(7.8)
        r2.font.color.rgb = c_text_secondary

    # =========================================================================
    # SLIDE 14: CONCLUSION & LIVE INTERACTIVE DEMO LINKS
    # =========================================================================
    add_interactive_header(s14, 13, "Conclusion & Interactive Live Demonstration")

    # Conclusion Card
    c_card_box = add_card(s14, Inches(0.8), Inches(1.35), Inches(11.733), Inches(2.4), title="FINAL PROJECT CONCLUSION", title_color=c_text_primary)
    tb_c = s14.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.333), Inches(1.8))
    tb_c.text_frame.word_wrap = True
    p_c = tb_c.text_frame.paragraphs[0]
    r_c = p_c.add_run()
    r_c.text = (
        "The complete engineering journey of Team 65 (Veltrox) from Phase 1 conceptualization to Phase 4 final cloud delivery "
        "demonstrates that digital privacy and physical confidentiality can be harmoniously unified without compromising usability.\n\n"
        "By fusing InspireFace AI Biometrics, Client-Side WebCrypto AES-GCM, Randomized Digital Steganography, and Ephemeral Forensic Self-Healing, "
        "Chatvia successfully eliminates the 'last-inch' presentation vulnerability. With 100% test verification and a live production cloud deployment, "
        "Chatvia stands as an innovative, robust, and battle-tested milestone in confidential communication."
    )
    r_c.font.size = Pt(9.5)
    r_c.font.color.rgb = c_text_secondary

    # Interactive Action Buttons Row
    links = [
        ("🚀 Launch Live Cloud Application", "Public Render Deployment", "https://chatvia-app-k7o7.onrender.com/", c_teal, c_teal_light, c_teal_dark),
        ("📂 GitHub Source Code Repository", "Full Source & History", "https://github.com/Kishan-Kumar-Hv/hmit", c_blue, c_blue_light, c_blue_dark),
        ("📄 Complete Project Documentation", "Phase 4 Master PDF Report", "Phase_4_Final_Project_Submission_Veltrox.pdf", c_emerald, c_emerald_light, c_emerald),
    ]

    for i, (l_title, l_sub, l_url, l_accent, l_bg, l_txt) in enumerate(links):
        lx = Inches(0.8) + i * Inches(3.98)
        l_btn = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, lx, Inches(3.95), Inches(3.78), Inches(1.35))
        l_btn.fill.solid()
        l_btn.fill.fore_color.rgb = l_bg
        l_btn.line.color.rgb = l_accent
        l_btn.line.width = Pt(1.5)

        tf_l = l_btn.text_frame
        tf_l.vertical_anchor = MSO_ANCHOR.MIDDLE
        p1 = tf_l.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        r1 = p1.add_run()
        r1.text = l_title + "\n"
        r1.font.bold = True
        r1.font.size = Pt(10.5)
        r1.font.color.rgb = c_text_primary
        r1.hyperlink.address = l_url

        p2 = tf_l.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = l_sub
        r2.font.bold = True
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = l_txt

    # Team Sign-off Card
    s_box = add_card(s14, Inches(0.8), Inches(5.5), Inches(11.733), Inches(1.4), border_color=c_teal)
    tb_s = s14.shapes.add_textbox(Inches(1.0), Inches(5.6), Inches(11.333), Inches(1.2))
    tb_s.text_frame.word_wrap = True
    p_s = tb_s.text_frame.paragraphs[0]
    r_s1 = p_s.add_run()
    r_s1.text = "SUBMITTED BY: TEAM 65 — VELTROX\n"
    r_s1.font.bold = True
    r_s1.font.size = Pt(11)
    r_s1.font.color.rgb = c_teal_dark

    r_s2 = p_s.add_run()
    r_s2.text = (
        "• Kishan Kumar H V (Lead & Cloud/DevOps Architect)    "
        "• Santhosh J R (Full-Stack & Biometrics Core Engineer)\n"
        "• Suhas B R (Security & Steganography Architect)       "
        "• Dhruva Kumar V M (Frontend Lead & UI/UX Specialist)\n"
        "HackMITTen 2026 | Cybersecurity & Applied AI Track | Final Phase 4 Submission Verified"
    )
    r_s2.font.size = Pt(8.5)
    r_s2.font.color.rgb = c_text_secondary

    prs.save(filename)
    print(f"Successfully generated light-mode interactive presentation deck: {filename}")

if __name__ == "__main__":
    output_file = "Phase_4_Final_Project_Presentation_Team65_Veltrox.pptx"
    if len(sys.argv) > 1:
        output_file = sys.argv[1]
    create_light_deck(output_file)
