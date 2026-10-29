import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_deck(filename="Phase_4_Final_Project_Presentation_Team65_Veltrox.pptx"):
    prs = Presentation()
    # 16:9 Widescreen format
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Palette
    c_bg = RGBColor(15, 23, 42)          # Slate 900
    c_card = RGBColor(30, 41, 59)        # Slate 800
    c_card_border = RGBColor(51, 65, 85) # Slate 700
    c_card_subtle = RGBColor(24, 33, 47) # Slate 850
    c_teal = RGBColor(13, 148, 136)      # Teal 600
    c_emerald = RGBColor(16, 185, 129)   # Emerald 500
    c_blue = RGBColor(59, 130, 246)      # Blue 500
    c_amber = RGBColor(245, 158, 11)     # Amber 500
    c_white = RGBColor(255, 255, 255)    # White
    c_slate200 = RGBColor(226, 232, 240)
    c_slate400 = RGBColor(148, 163, 184)
    c_slate500 = RGBColor(100, 116, 139)
    c_muted = c_slate500

    nav_items = [
        ("Overview", 2),
        ("Problem", 3),
        ("Solution", 4),
        ("Methodology", 5),
        ("Architecture", 6),
        ("Tech Stack", 7),
        ("Features", 8),
        ("Testing", 9),
        ("Screenshots", 10),
        ("Challenges", 12),
        ("Results & Future", 13),
    ]

    def add_background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = c_bg
        bg.line.fill.background()
        return bg

    def add_top_bar(slide, active_title, category="PHASE 4: FINAL PROJECT PRESENTATION"):
        # Category breadcrumb
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.733), Inches(0.4))
        p = tb.text_frame.paragraphs[0]
        r1 = p.add_run()
        r1.text = f"TEAM 65 (VELTROX)  |  HACKMITTEN 2026  •  "
        r1.font.bold = True
        r1.font.size = Pt(9)
        r1.font.color.rgb = c_teal

        r2 = p.add_run()
        r2.text = category
        r2.font.bold = False
        r2.font.size = Pt(9)
        r2.font.color.rgb = c_slate400

        # Slide Title
        tb2 = slide.shapes.add_textbox(Inches(0.8), Inches(0.65), Inches(9.5), Inches(0.6))
        p2 = tb2.text_frame.paragraphs[0]
        r3 = p2.add_run()
        r3.text = active_title
        r3.font.bold = True
        r3.font.size = Pt(20)
        r3.font.color.rgb = c_white

        # Interactive Navigation Buttons on Top Right
        btn_w = Inches(1.35)
        btn_h = Inches(0.36)
        btn_y = Inches(0.65)
        
        # Live Demo Link Button
        b1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.6), btn_y, btn_w, btn_h)
        b1.fill.solid()
        b1.fill.fore_color.rgb = c_teal
        b1.line.color.rgb = c_emerald
        b1.line.width = Pt(1)
        tf1 = b1.text_frame
        tf1.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_b1 = tf1.paragraphs[0]
        p_b1.alignment = PP_ALIGN.CENTER
        r_b1 = p_b1.add_run()
        r_b1.text = "⚡ Live Demo"
        r_b1.font.bold = True
        r_b1.font.size = Pt(9)
        r_b1.font.color.rgb = c_white
        r_b1.hyperlink.address = "https://chatvia-app-k7o7.onrender.com/"

        # GitHub Repo Link Button
        b2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(12.05), btn_y, Inches(0.9), btn_h)
        b2.fill.solid()
        b2.fill.fore_color.rgb = c_card
        b2.line.color.rgb = c_card_border
        b2.line.width = Pt(1)
        tf2 = b2.text_frame
        tf2.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_b2 = tf2.paragraphs[0]
        p_b2.alignment = PP_ALIGN.CENTER
        r_b2 = p_b2.add_run()
        r_b2.text = "GitHub"
        r_b2.font.bold = True
        r_b2.font.size = Pt(9)
        r_b2.font.color.rgb = c_slate200
        r_b2.hyperlink.address = "https://github.com/Kishan-Kumar-Hv/hmit"

    def add_card(slide, left, top, width, height, title="", border_color=c_card_border, bg_color=c_card):
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        box.fill.solid()
        box.fill.fore_color.rgb = bg_color
        box.line.color.rgb = border_color
        box.line.width = Pt(1)
        
        if title:
            tb = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), width - Inches(0.4), Inches(0.4))
            p = tb.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = title
            r.font.bold = True
            r.font.size = Pt(12)
            r.font.color.rgb = c_teal
        return box

    # =========================================================================
    # SLIDE 1: TITLE & TEAM DETAILS
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    add_background(s1)

    # Accent decorative top strip
    strip = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.1))
    strip.fill.solid()
    strip.fill.fore_color.rgb = c_teal
    strip.line.fill.background()

    # Track Badge
    badge = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.8), Inches(3.6), Inches(0.4))
    badge.fill.solid()
    badge.fill.fore_color.rgb = c_card
    badge.line.color.rgb = c_teal
    tf_badge = badge.text_frame
    tf_badge.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_badge = tf_badge.paragraphs[0]
    p_badge.alignment = PP_ALIGN.CENTER
    r_badge = p_badge.add_run()
    r_badge.text = "HACKMITTEN 2026 • CYBERSECURITY & AI"
    r_badge.font.bold = True
    r_badge.font.size = Pt(9)
    r_badge.font.color.rgb = c_teal

    # Project Title
    t_title = s1.shapes.add_textbox(Inches(0.8), Inches(1.3), Inches(11.733), Inches(1.4))
    p_t = t_title.text_frame.paragraphs[0]
    r_t = p_t.add_run()
    r_t.text = "CHATVIA"
    r_t.font.bold = True
    r_t.font.size = Pt(44)
    r_t.font.color.rgb = c_white

    p_sub = t_title.text_frame.add_paragraph()
    r_sub = p_sub.add_run()
    r_sub.text = "Next-Generation Confidential Messaging with AI Biometrics & Client-Side Steganography"
    r_sub.font.bold = True
    r_sub.font.size = Pt(17)
    r_sub.font.color.rgb = c_teal

    p_desc = t_title.text_frame.add_paragraph()
    r_desc = p_desc.add_run()
    r_desc.text = "Eliminating the 'Last-Inch' Presentation Vulnerability through Continuous Facial Verification & Ephemeral Forensic Self-Healing"
    r_desc.font.size = Pt(11.5)
    r_desc.font.color.rgb = c_slate400

    # Team 65 / Veltrox Card
    add_card(s1, Inches(0.8), Inches(3.3), Inches(11.733), Inches(3.2), title="TEAM 65 — VELTROX | ROLES & CONTRIBUTIONS")

    members = [
        ("Kishan Kumar H V", "Team Lead & Cloud/DevOps", "Production PaaS deployment on Render, Express 5 reverse proxy, Supabase SSL pooler engineering, system testing & documentation."),
        ("Santhosh J R", "Full-Stack & Biometrics", "InspireFace AI microservice integration, 512-d feature extraction, WebSocket messaging hub, active surveillance loop & warning banners."),
        ("Suhas B R", "Security & Steganography", "Client WebCrypto AES-GCM pipeline, friendship secret agreement with PBKDF2 salted hash, IndexedDB keys, and randomized LSB engine."),
        ("Dhruva Kumar V M", "Frontend Lead & UI/UX", "Modern React 19 SPA architecture, responsive Tailwind dark/light theme, WebRTC camera lifecycle, file upload streaming & QA."),
    ]

    col_w = Inches(2.78)
    for i, (m_name, m_role, m_desc) in enumerate(members):
        m_x = Inches(0.95) + i * Inches(2.88)
        m_box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, m_x, Inches(3.9), col_w, Inches(2.35))
        m_box.fill.solid()
        m_box.fill.fore_color.rgb = c_card_subtle
        m_box.line.color.rgb = c_card_border
        
        tf_m = m_box.text_frame
        tf_m.word_wrap = True
        
        p1 = tf_m.paragraphs[0]
        r_name = p1.add_run()
        r_name.text = m_name
        r_name.font.bold = True
        r_name.font.size = Pt(11)
        r_name.font.color.rgb = c_white

        p2 = tf_m.add_paragraph()
        r_role = p2.add_run()
        r_role.text = m_role
        r_role.font.bold = True
        r_role.font.size = Pt(9)
        r_role.font.color.rgb = c_teal

        p3 = tf_m.add_paragraph()
        r_desc = p3.add_run()
        r_desc.text = m_desc
        r_desc.font.size = Pt(8)
        r_desc.font.color.rgb = c_slate400

    # Bottom Quick Interactive Links
    lbl = s1.shapes.add_textbox(Inches(0.8), Inches(6.75), Inches(11.733), Inches(0.4))
    p_lnk = lbl.text_frame.paragraphs[0]
    r_l1 = p_lnk.add_run()
    r_l1.text = "⚡ Live Cloud Application: "
    r_l1.font.bold = True
    r_l1.font.size = Pt(10)
    r_l1.font.color.rgb = c_slate400

    r_l2 = p_lnk.add_run()
    r_l2.text = "https://chatvia-app-k7o7.onrender.com/   "
    r_l2.font.bold = True
    r_l2.font.size = Pt(10)
    r_l2.font.color.rgb = c_blue
    r_l2.hyperlink.address = "https://chatvia-app-k7o7.onrender.com/"

    r_l3 = p_lnk.add_run()
    r_l3.text = "•   GitHub Source: "
    r_l3.font.bold = True
    r_l3.font.size = Pt(10)
    r_l3.font.color.rgb = c_slate400

    r_l4 = p_lnk.add_run()
    r_l4.text = "https://github.com/Kishan-Kumar-Hv/hmit"
    r_l4.font.bold = True
    r_l4.font.size = Pt(10)
    r_l4.font.color.rgb = c_blue
    r_l4.hyperlink.address = "https://github.com/Kishan-Kumar-Hv/hmit"

    # =========================================================================
    # SLIDE 2: EXECUTIVE SUMMARY & PROJECT JOURNEY
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_background(s2)
    add_top_bar(s2, "Executive Summary: The Complete Journey (Phases 1 → 4)")

    phases = [
        ("Phase 1: Initial Analysis", "IDEATION & CARDBOARD", c_teal, [
            "Brainstormed physical snooping threat models",
            "Conceived 3-tier defense on physical cardboard",
            "Rejected heavy client WASM in favor of backend AI",
            "Established privacy rule: Zero raw facial photos saved",
            "Designed randomized PRNG LSB steganography"
        ]),
        ("Phase 2: Implementation", "ENGINEERING SPRINTS", c_blue, [
            "Sprint 1: Auth & InspireFace 512-d embeddings",
            "Sprint 2: WebSocket instant chat & 25MB attachments",
            "Sprint 3: WebCrypto AES-GCM + IndexedDB key cache",
            "Sprint 4: Active face surveillance loop & timeout aborts",
            "Unified single-port Express 5 SPA reverse proxy"
        ]),
        ("Phase 3: Testing & Security", "VERIFICATION & QA", c_amber, [
            "Executed 20 rigorous test cases (100% Pass)",
            "Validated single vs multi-face gating",
            "Tested shoulder-surfing instant auto-concealment",
            "Verified 25MB attachment quota boundary",
            "Proved database ephemeral self-healing cleanup"
        ]),
        ("Phase 4: Cloud Delivery", "PRODUCTION LAUNCH", c_emerald, [
            "Production PaaS deployment on Render",
            "Supabase PostgreSQL pooler integration with SSL",
            "Automated CI/CD build from GitHub main branch",
            "Sub-80ms message delivery; sub-70ms face verification",
            "Comprehensive technical documentation & presentation"
        ]),
    ]

    for i, (p_title, p_sub, p_color, p_bullets) in enumerate(phases):
        card_x = Inches(0.8) + i * Inches(2.98)
        c_box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, card_x, Inches(1.5), Inches(2.82), Inches(5.3))
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = c_card
        c_box.line.color.rgb = p_color
        c_box.line.width = Pt(1.5)

        tf = c_box.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        r1 = p.add_run()
        r1.text = p_title + "\n"
        r1.font.bold = True
        r1.font.size = Pt(13)
        r1.font.color.rgb = c_white

        r2 = p.add_run()
        r2.text = p_sub + "\n\n"
        r2.font.bold = True
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = p_color

        for b in p_bullets:
            pb = tf.add_paragraph()
            pb.space_after = Pt(8)
            rb = pb.add_run()
            rb.text = "• " + b
            rb.font.size = Pt(9)
            rb.font.color.rgb = c_slate200

    # =========================================================================
    # SLIDE 3: PROBLEM STATEMENT & THREAT ANALYSIS
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_background(s3)
    add_top_bar(s3, "Problem Statement & The 'Last-Inch' Vulnerability Gap")

    # Problem Statement Callout Card
    ps_card = add_card(s3, Inches(0.8), Inches(1.4), Inches(11.733), Inches(1.3), title="ORIGINAL PROBLEM STATEMENT ASSIGNED TO TEAM 65")
    tf_ps = ps_card.text_frame
    p_ps = tf_ps.add_paragraph()
    p_ps.space_before = Pt(16)
    r_ps = p_ps.add_run()
    r_ps.text = '"Design and engineer a real-time secure communication system that eliminates vulnerabilities in both data transmission and physical message presentation. The application must guarantee that confidential exchanges remain fully protected from over-the-shoulder eavesdropping, unauthorized physical device access, and persistent forensic extraction, while providing zero-knowledge verifiable privacy and high usability."'
    r_ps.font.italic = True
    r_ps.font.size = Pt(10)
    r_ps.font.color.rgb = c_slate200

    # 4 Vulnerability Cards
    gaps = [
        ("1. The 'Last-Inch' Presentation Gap", "Conventional E2EE (Signal, WhatsApp) encrypts data in flight, but once decrypted onto the screen, it is fully vulnerable to shoulder surfers, bystanders, and unauthorized onlookers in coffee shops, transit, or offices.", c_amber),
        ("2. Static vs Continuous Auth", "Face ID and PIN locks authenticate identity only ONCE when opening the application. If the unlocked phone is set down, passed to a colleague, or seized, any person can scroll through confidential chats unchecked.", c_amber),
        ("3. Persistent Forensic Media Residue", "Standard apps persist decrypted media to disk, gallery caches, and SQLite databases. Even after 'deleting' messages, forensic tools easily extract image artifacts from operating system flash blocks.", c_amber),
        ("4. Traffic Observability & Deniability", "Sending raw encrypted payloads alerts network firewalls and eavesdroppers that sensitive data is in transit. Users lack plausible deniability without covert carrier concealment.", c_amber)
    ]

    for i, (g_title, g_desc, g_color) in enumerate(gaps):
        gx = Inches(0.8) + (i % 2) * Inches(5.95)
        gy = Inches(2.95) + (i // 2) * Inches(2.05)
        add_card(s3, gx, gy, Inches(5.78), Inches(1.85), title=g_title, border_color=g_color)
        
        tb = s3.shapes.add_textbox(gx + Inches(0.2), gy + Inches(0.45), Inches(5.38), Inches(1.2))
        tb.text_frame.word_wrap = True
        p = tb.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = g_desc
        r.font.size = Pt(9.5)
        r.font.color.rgb = c_slate200

    # =========================================================================
    # SLIDE 4: PROPOSED SOLUTION & CORE OBJECTIVES
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_background(s4)
    add_top_bar(s4, "Proposed Solution: Multi-Layered Zero-Trust Architecture")

    pillars = [
        ("AI Biometric Active Shield", "Continuous Facial Verification", c_teal, [
            "InspireFace Deep Learning SDK",
            "512-dimensional facial embedding vectors",
            "Cosine similarity thresholding (>= 0.70)",
            "Enforces exactly 1 face present at all times",
            "Continuous 1.5s background surveillance loop",
            "Auto-conceals immediately if user turns away or bystander detected"
        ]),
        ("Randomized Steganography", "Plausible Deniability Transport", c_blue, [
            "Disguises secret text inside innocent PNGs",
            "24-byte header with CVS1 magic tag",
            "16-byte cryptographically secure PRNG seed",
            "Uniform pseudo-random pixel channel scattering",
            "Zero perceptible visual distortion in images",
            "Protects against network payload inspection"
        ]),
        ("Client-Side WebCrypto", "Zero-Knowledge Key Agreement", c_emerald, [
            "Out-of-band mutual friendship passphrase",
            "Client PBKDF2-SHA256 (100k rounds) derivation",
            "AES-GCM-256 authenticated encryption",
            "Keys marked non-exportable in IndexedDB",
            "Server receives only salted verifier hash",
            "Server never has access to plaintext or keys"
        ]),
        ("Forensic Self-Healing", "Ephemeral Database Destruction", c_amber, [
            "One-time secure message consumption",
            "Atomic PostgreSQL database transaction",
            "Restores clean original carrier PNG to attachments",
            "Permanently nullifies original_png backup",
            "Eliminates forensic stego traces on disk",
            "Guarantees forward secrecy post-viewing"
        ]),
    ]

    for i, (p_title, p_sub, p_color, p_items) in enumerate(pillars):
        px = Inches(0.8) + i * Inches(2.98)
        p_box = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px, Inches(1.4), Inches(2.82), Inches(5.4))
        p_box.fill.solid()
        p_box.fill.fore_color.rgb = c_card
        p_box.line.color.rgb = p_color
        p_box.line.width = Pt(1.5)

        tf = p_box.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        r1 = p.add_run()
        r1.text = p_title + "\n"
        r1.font.bold = True
        r1.font.size = Pt(12)
        r1.font.color.rgb = c_white

        r2 = p.add_run()
        r2.text = p_sub + "\n\n"
        r2.font.bold = True
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = p_color

        for item in p_items:
            pi = tf.add_paragraph()
            pi.space_after = Pt(7)
            ri = pi.add_run()
            ri.text = "• " + item
            ri.font.size = Pt(8.5)
            ri.font.color.rgb = c_slate200

    # =========================================================================
    # SLIDE 5: SYSTEM METHODOLOGY & WORKFLOW
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_background(s5)
    add_top_bar(s5, "Methodology: End-to-End System Workflow (8 Phases)")

    steps = [
        ("01", "Registration & Consent", "Camera captures face; InspireFace validates exactly 1 face; 512-d template encrypted via AES-256-GCM; raw image purged."),
        ("02", "Friendship Handshake", "Users exchange unique 32-character User ID codes; friend request accepted; mutual privacy verified."),
        ("03", "Key Derivation", "Both users input shared passphrase; browser WebCrypto API derives non-exportable AES-GCM key into IndexedDB."),
        ("04", "Steganographic Send", "Sender encrypts secret text; seeds PRNG across PNG pixel channels; stego PNG and recovery copy sent to server."),
        ("05", "Real-Time Delivery", "Express backend persists BYTEA to PostgreSQL; broadcasts event over WebSocket (ws://); recipient receives message card."),
        ("06", "Biometric Unlock", "Recipient clicks decrypt; camera captures face; FastAPI AI verifies cosine similarity >= 0.70; secret revealed."),
        ("07", "Active Surveillance", "Continuous 1.5s verification loop runs with 2s timeout; if face count != 1 or face mismatch, text is instantly concealed."),
        ("08", "Forensic Self-Healing", "Message consumed; database atomically restores clean original PNG to attachment and deletes recovery copy."),
    ]

    for i, (num, s_title, s_desc) in enumerate(steps):
        col = i % 4
        row = i // 4
        sx = Inches(0.8) + col * Inches(2.98)
        sy = Inches(1.5) + row * Inches(2.65)

        s_box = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx, sy, Inches(2.82), Inches(2.45))
        s_box.fill.solid()
        s_box.fill.fore_color.rgb = c_card
        s_box.line.color.rgb = c_teal if row == 0 else c_blue
        s_box.line.width = Pt(1)

        tf = s_box.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        r_num = p.add_run()
        r_num.text = f"PHASE {num}\n"
        r_num.font.bold = True
        r_num.font.size = Pt(9)
        r_num.font.color.rgb = c_teal if row == 0 else c_blue

        r_title = p.add_run()
        r_title.text = s_title + "\n\n"
        r_title.font.bold = True
        r_title.font.size = Pt(11)
        r_title.font.color.rgb = c_white

        p_desc = tf.add_paragraph()
        r_desc = p_desc.add_run()
        r_desc.text = s_desc
        r_desc.font.size = Pt(8.5)
        r_desc.font.color.rgb = c_slate200

    # =========================================================================
    # SLIDE 6: SYSTEM ARCHITECTURE
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_background(s6)
    add_top_bar(s6, "System Architecture: Decoupled 4-Tier Cloud Architecture")

    # Tier 1: Client
    t1 = add_card(s6, Inches(0.8), Inches(1.4), Inches(11.733), Inches(1.2), title="TIER 1: PRESENTATION LAYER (React 19 + TypeScript + Vite + Tailwind CSS)", border_color=c_teal)
    tb1 = s6.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.333), Inches(0.7))
    tb1.text_frame.word_wrap = True
    p1 = tb1.text_frame.paragraphs[0]
    p1.add_run().text = "• WebCrypto API (AES-GCM-256) | Non-exportable IndexedDB Key Persistence | WebRTC MediaStream Camera Acquisition\n• Real-Time WebSocket Client (ws://) | Active Viewer 1.5s Continuous Surveillance Loop | AbortSignal 2s Timeout Protection"
    p1.runs[0].font.size = Pt(9)
    p1.runs[0].font.color.rgb = c_slate200

    # Connector arrow 1
    arrow1 = s6.shapes.add_textbox(Inches(5.0), Inches(2.6), Inches(3.333), Inches(0.35))
    p_a1 = arrow1.text_frame.paragraphs[0]
    p_a1.alignment = PP_ALIGN.CENTER
    r_a1 = p_a1.add_run()
    r_a1.text = "▼ RESTful JSON APIs & Bi-Directional WebSockets ▼"
    r_a1.font.bold = True
    r_a1.font.size = Pt(8.5)
    r_a1.font.color.rgb = c_teal

    # Tier 2: Gateway
    t2 = add_card(s6, Inches(0.8), Inches(2.95), Inches(11.733), Inches(1.2), title="TIER 2: GATEWAY & APPLICATION SERVER (Node.js + Express 5)", border_color=c_blue)
    tb2 = s6.shapes.add_textbox(Inches(1.0), Inches(3.35), Inches(11.333), Inches(0.7))
    tb2.text_frame.word_wrap = True
    p2 = tb2.text_frame.paragraphs[0]
    p2.add_run().text = "• Single-Port Reverse Proxy & Vite SPA Static Server | JWT Auth & bcrypt Password Verification | Multer 25MB File Streaming\n• Master AES-256-GCM Biometric Template Encryptor | WebSocket Broadcast State Hub | PostgreSQL Connection Pool with SSL"
    p2.runs[0].font.size = Pt(9)
    p2.runs[0].font.color.rgb = c_slate200

    # Connector arrow 2 (Split)
    arrow2a = s6.shapes.add_textbox(Inches(2.0), Inches(4.15), Inches(3.0), Inches(0.35))
    p_a2a = arrow2a.text_frame.paragraphs[0]
    p_a2a.alignment = PP_ALIGN.CENTER
    r_a2a = p_a2a.add_run()
    r_a2a.text = "▼ Internal HTTP (:8001) ▼"
    r_a2a.font.bold = True
    r_a2a.font.size = Pt(8.5)
    r_a2a.font.color.rgb = c_muted

    arrow2b = s6.shapes.add_textbox(Inches(8.0), Inches(4.15), Inches(3.0), Inches(0.35))
    p_a2b = arrow2b.text_frame.paragraphs[0]
    p_a2b.alignment = PP_ALIGN.CENTER
    r_a2b = p_a2b.add_run()
    r_a2b.text = "▼ Encrypted SSL TCP (:5432) ▼"
    r_a2b.font.bold = True
    r_a2b.font.size = Pt(8.5)
    r_a2b.font.color.rgb = c_muted

    # Tier 3: AI Microservice
    t3 = add_card(s6, Inches(0.8), Inches(4.5), Inches(5.75), Inches(2.3), title="TIER 3: AI & STEGO ENGINE (Python FastAPI)", border_color=c_emerald)
    tb3 = s6.shapes.add_textbox(Inches(1.0), Inches(4.9), Inches(5.35), Inches(1.8))
    tb3.text_frame.word_wrap = True
    p3 = tb3.text_frame.paragraphs[0]
    p3.add_run().text = "• InspireFace SDK ('Pikachu' Model)\n• OpenCV Face Detection & Alignment\n• 512-d Embedding Vector Extraction\n• Cosine Similarity Verification (>= 0.70)\n• Thread-Safe Session Locking (_session_lock)\n• Randomized PRNG LSB Steganography Engine"
    p3.runs[0].font.size = Pt(8.5)
    p3.runs[0].font.color.rgb = c_slate200

    # Tier 4: Database
    t4 = add_card(s6, Inches(6.78), Inches(4.5), Inches(5.75), Inches(2.3), title="TIER 4: PERSISTENCE (PostgreSQL 16 Supabase)", border_color=c_amber)
    tb4 = s6.shapes.add_textbox(Inches(6.98), Inches(4.9), Inches(5.35), Inches(1.8))
    tb4.text_frame.word_wrap = True
    p4 = tb4.text_frame.paragraphs[0]
    p4.add_run().text = "• users: Profiles & bcrypt password hashes\n• face_biometrics: AES-256-GCM ciphertexts + nonces\n• friend_requests & friendship_secrets: Salted hashes\n• messages & attachments: BYTEA binary blocks (25MB)\n• stego_messages: Ephemeral carrier recovery copies\n• Automated atomic self-healing triggers"
    p4.runs[0].font.size = Pt(8.5)
    p4.runs[0].font.color.rgb = c_slate200

    # =========================================================================
    # SLIDE 7: TECHNOLOGIES & TOOLS USED
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_background(s7)
    add_top_bar(s7, "Technologies, Frameworks & Tooling Stack")

    tech_domains = [
        ("Frontend & Client Security", c_teal, [
            ("React 19 & TypeScript", "Modular component state architecture & type safety"),
            ("Vite 8", "High-speed modern frontend bundler & asset compiler"),
            ("Tailwind CSS", "Responsive dark/light themed utility-first design"),
            ("Web Crypto API", "Client-side PBKDF2 key derivation & AES-GCM cipher"),
            ("IndexedDB", "Secure browser-isolated non-exportable key storage"),
            ("HTML5 WebRTC", "Real-time webcam MediaStream acquisition & frame sampling")
        ]),
        ("Backend & Real-Time Gateway", c_blue, [
            ("Node.js & Express 5", "Unified RESTful API server & static SPA host"),
            ("WebSockets (`ws`)", "Bi-directional real-time message & status hub"),
            ("JSON Web Tokens (JWT)", "Stateless secure bearer session authentication"),
            ("bcrypt", "Cryptographic password hashing (10 salt rounds)"),
            ("Multer", "Memory-buffered multipart attachment handling (25MB)"),
            ("node-postgres (`pg`)", "High-performance connection pool with dynamic SSL")
        ]),
        ("AI Vision & Steganography", c_emerald, [
            ("Python 3.12 & FastAPI", "High-throughput asynchronous AI microservice"),
            ("InspireFace SDK", "Deep face detection & 512-d feature extraction"),
            ("OpenCV (`cv2`)", "Raw frame matrix decoding, resizing & alignment"),
            ("NumPy 2.4", "Vectorized array manipulation & PRNG bit scattering"),
            ("Pillow (`PIL`)", "Lossless RGB PNG encoding/decoding"),
            ("Python Secrets", "Cryptographically secure 16-byte PRNG token generation")
        ]),
        ("Database, Cloud & DevOps", c_amber, [
            ("PostgreSQL 16", "ACID compliant relational storage with BYTEA support"),
            ("Supabase Managed DB", "Cloud database with SSL transaction pooler"),
            ("Render PaaS", "Unified single-domain container cloud deployment"),
            ("Git & GitHub", "Distributed version control & CI/CD deployment"),
            ("npm & package.json", "Unified root monorepo build pipeline orchestration"),
            ("PyMuPDF & ReportLab", "Automated PDF documentation & test report generation")
        ]),
    ]

    for i, (dom_name, dom_color, dom_items) in enumerate(tech_domains):
        dx = Inches(0.8) + i * Inches(2.98)
        d_box = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, dx, Inches(1.4), Inches(2.82), Inches(5.4))
        d_box.fill.solid()
        d_box.fill.fore_color.rgb = c_card
        d_box.line.color.rgb = dom_color
        d_box.line.width = Pt(1.5)

        tf = d_box.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = dom_name + "\n\n"
        r.font.bold = True
        r.font.size = Pt(11.5)
        r.font.color.rgb = c_white

        for tech, desc in dom_items:
            pt = tf.add_paragraph()
            pt.space_after = Pt(6)
            rt = pt.add_run()
            rt.text = tech + "\n"
            rt.font.bold = True
            rt.font.size = Pt(8.5)
            rt.font.color.rgb = dom_color

            rd = pt.add_run()
            rd.text = desc
            rd.font.size = Pt(7.5)
            rd.font.color.rgb = c_slate200

    # =========================================================================
    # SLIDE 8: PHASE 2 IMPLEMENTATION & MAJOR FEATURES
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_background(s8)
    add_top_bar(s8, "Phase 2 Implementation: 5 Core Innovations Developed")

    innovations = [
        ("Feature 1: Zero-Knowledge Biometric Enrollment", "Mandatory facial registration captures exactly 1 face using InspireFace. Extracts 512-d normalized floating point embeddings. Node.js backend encrypts vector with AES-256-GCM using FACE_ENCRYPTION_KEY before database insertion. The raw user photo is purged from memory immediately—never stored on disk.", c_teal),
        ("Feature 2: Randomized LSB Steganography Engine", "Disguises encrypted secret text into lossless PNG carriers. Writes a 24-byte header (CVS1 magic tag, 16-byte random PRNG seed, 4-byte payload size) into initial pixels, then scatters encrypted bits randomly across RGB channels via NumPy default_rng. Resulting image exhibits zero perceptible visual degradation.", c_blue),
        ("Feature 3: Active Face Guardian Surveillance Loop", "Decryption activates recipient webcam. The AI verifies identity against enrolled template (cosine >= 0.70) and guarantees exactly 1 face. A continuous surveillance loop checks the viewer every 1.5s with a 2s AbortSignal timeout. If the user turns away or a second face appears, text is auto-concealed.", c_emerald),
        ("Feature 4: Single-Port Cloud Bundling on Render", "Cloud PaaS provides only one public port ($PORT). Team Veltrox engineered a root build script compiling Vite SPA and Express server into a unified artifact. Express 5 acts as reverse proxy, serving React statically from dist/ and handling catch-all SPA routing seamlessly.", c_amber),
        ("Feature 5: Forensic Database Self-Healing", "Traditional apps leave persistent traces of decrypted media. Chatvia implements atomic self-healing: upon single-view decryption, PostgreSQL updates attachments with the original clean PNG and nullifies original_png, removing any forensic trace of steganography.", c_teal),
    ]

    for i, (f_title, f_desc, f_color) in enumerate(innovations):
        fy = Inches(1.4) + i * Inches(1.08)
        f_card = add_card(s8, Inches(0.8), fy, Inches(11.733), Inches(0.98), title=f_title, border_color=f_color)
        
        tb = s8.shapes.add_textbox(Inches(1.0), fy + Inches(0.35), Inches(11.333), Inches(0.55))
        tb.text_frame.word_wrap = True
        p = tb.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = f_desc
        r.font.size = Pt(8.5)
        r.font.color.rgb = c_slate200

    # =========================================================================
    # SLIDE 9: PHASE 3 TESTING & VERIFICATION (100% PASS)
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_background(s9)
    add_top_bar(s9, "Phase 3 Testing: 20 Test Cases Executed (100% Pass Rate)")

    # Test Metrics Row
    metrics = [
        ("Total Test Cases", "20", c_blue),
        ("Test Execution Pass", "100%", c_emerald),
        ("Biometric Latency", "< 70 ms", c_teal),
        ("Message Delivery", "< 80 ms", c_teal),
        ("Max File Upload", "25 MB", c_amber),
        ("Timeout Abort Limit", "2,000 ms", c_amber),
    ]

    for i, (m_lbl, m_val, m_col) in enumerate(metrics):
        mx = Inches(0.8) + i * Inches(1.98)
        m_box = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, mx, Inches(1.4), Inches(1.82), Inches(0.9))
        m_box.fill.solid()
        m_box.fill.fore_color.rgb = c_card
        m_box.line.color.rgb = m_col
        m_box.line.width = Pt(1)

        tf = m_box.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p1 = tf.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        r1 = p1.add_run()
        r1.text = m_val + "\n"
        r1.font.bold = True
        r1.font.size = Pt(16)
        r1.font.color.rgb = m_col

        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = m_lbl
        r2.font.size = Pt(7.5)
        r2.font.color.rgb = c_slate400

    # Test Case Table
    t_shape = s9.shapes.add_table(7, 5, Inches(0.8), Inches(2.45), Inches(11.733), Inches(4.5))
    table = t_shape.table
    table.columns[0].width = Inches(1.2)
    table.columns[1].width = Inches(1.8)
    table.columns[2].width = Inches(3.2)
    table.columns[3].width = Inches(4.333)
    table.columns[4].width = Inches(1.2)

    headers = ["Test ID", "Module", "Condition / Input Tested", "Expected & Verified Outcome", "Status"]
    for col_idx, h_text in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = c_teal
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = h_text
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = c_white

    sample_tests = [
        ("TC-REG-02", "Biometric Enrollment", "User attempts registration with 2 faces visible", "HTTP 422 rejected; error 'Capture exactly one face, clearly visible'; blocked", "PASS (100%)"),
        ("TC-KEY-01", "Friendship Key Derivation", "Both friends input identical shared passphrase", "PBKDF2 salted hash verified; non-exportable AES-GCM key cached in IndexedDB", "PASS (100%)"),
        ("TC-STEG-01", "Randomized LSB Stego", "500-char private text embedded into 1.2MB PNG", "Header CVS1 written; payload scattered via PRNG; zero visual distortion", "PASS (100%)"),
        ("TC-BIO-03", "Shoulder Surfer Detection", "Second person appears in webcam while text is open", "Surveillance detects face_count == 2; text auto-concealed immediately", "PASS (100%)"),
        ("TC-BIO-05", "Timeout Protection", "Simulated network latency > 2,000ms during check", "AbortSignal.timeout(2000) cancels request; text concealed safely", "PASS (100%)"),
        ("TC-HEAL-01", "Forensic Self-Healing", "Message decrypted and successfully viewed once", "Database restores original clean PNG; recovery copy permanently nullified", "PASS (100%)"),
    ]

    for row_idx, row_data in enumerate(sample_tests, start=1):
        for col_idx, text in enumerate(row_data):
            cell = table.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = c_card if row_idx % 2 == 1 else c_card_subtle
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
                r.font.color.rgb = c_white
            else:
                r.font.color.rgb = c_slate200

    # =========================================================================
    # SLIDE 10: APPLICATION SCREENSHOTS - DESKTOP DASHBOARD
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_background(s10)
    add_top_bar(s10, "Application Visual Proof: Desktop Chatvia Dashboard")

    img1_path = "Screenshots/Screenshot 2026-10-29 034058.png"
    if os.path.exists(img1_path):
        # Insert screenshot with border card
        add_card(s10, Inches(0.8), Inches(1.4), Inches(8.5), Inches(5.4), border_color=c_teal)
        s10.shapes.add_picture(img1_path, Inches(0.9), Inches(1.5), Inches(8.3), Inches(5.2))

        # Annotations Card on the Right
        add_card(s10, Inches(9.5), Inches(1.4), Inches(3.033), Inches(5.4), title="KEY CAPABILITIES")
        tb = s10.shapes.add_textbox(Inches(9.65), Inches(1.85), Inches(2.733), Inches(4.8))
        tb.text_frame.word_wrap = True
        p = tb.text_frame.paragraphs[0]
        
        callouts = [
            ("Active Contact Directory", "Bilateral friend relationships using unique User ID codes. Displays online status in real-time."),
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
            r1.font.color.rgb = c_teal

            r2 = p_c.add_run()
            r2.text = c_text
            r2.font.size = Pt(7.5)
            r2.font.color.rgb = c_slate200

    # =========================================================================
    # SLIDE 11: APPLICATION SCREENSHOTS - MOBILE & BIOMETRICS
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    add_background(s11)
    add_top_bar(s11, "Application Visual Proof: Biometric Enrollment & Stego Sender")

    img2_path = "Screenshots/Screenshot 2026-10-29 034122.png"
    img3_path = "Screenshots/Screenshot 2026-10-29 034136.png"

    if os.path.exists(img2_path) and os.path.exists(img3_path):
        # Image 2 Box
        add_card(s11, Inches(0.8), Inches(1.4), Inches(3.4), Inches(5.4), title="BIOMETRIC ENROLLMENT", border_color=c_teal)
        s11.shapes.add_picture(img2_path, Inches(0.95), Inches(1.85), Inches(3.1), Inches(4.8))

        # Image 3 Box
        add_card(s11, Inches(4.4), Inches(1.4), Inches(3.4), Inches(5.4), title="PRIVATE STEGO MESSAGE", border_color=c_blue)
        s11.shapes.add_picture(img3_path, Inches(4.55), Inches(1.85), Inches(3.1), Inches(4.8))

        # Walkthrough Card on Right
        add_card(s11, Inches(8.0), Inches(1.4), Inches(4.533), Inches(5.4), title="MOBILE & SECURITY WALKTHROUGH")
        tb = s11.shapes.add_textbox(Inches(8.2), Inches(1.85), Inches(4.133), Inches(4.8))
        tb.text_frame.word_wrap = True
        
        walkthrough = [
            ("Mandatory Biometric Registration (Figure 2)", "Users cannot sign up without explicit camera consent and single-face capture. InspireFace validates alignment, eliminates multi-face fraud, and outputs a 512-d embedding. The raw photograph is instantly destroyed."),
            ("Master Key Encryption Before Persistence", "Biometric templates are encrypted with AES-256-GCM via the backend's secret master key prior to PostgreSQL storage, preventing data leakage in database dumps."),
            ("Private Image Message Composition (Figure 3)", "Users toggle 'Private image message' in the chat composer. An innocent carrier PNG is selected, and secret text is typed. The client encrypts the text with WebCrypto before sending to the stego engine."),
            ("Friendship Key Management & Rotation", "Friends agree on a shared passphrase once. Derived keys persist in IndexedDB without ever being sent across the wire. Key rotation is blocked if unread stego messages exist.")
        ]
        for w_title, w_desc in walkthrough:
            p_w = tb.text_frame.add_paragraph()
            p_w.space_after = Pt(10)
            r1 = p_w.add_run()
            r1.text = "• " + w_title + ":\n"
            r1.font.bold = True
            r1.font.size = Pt(9)
            r1.font.color.rgb = c_teal

            r2 = p_w.add_run()
            r2.text = w_desc
            r2.font.size = Pt(8)
            r2.font.color.rgb = c_slate200

    # =========================================================================
    # SLIDE 12: CHALLENGES FACED & SOLUTIONS
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    add_background(s12)
    add_top_bar(s12, "Challenges Faced & Engineering Solutions")

    challenges = [
        ("1. Render Single-Port Limitation", "Cloud PaaS provides only one public web port ($PORT), but the architecture requires a React SPA, Express REST API, and WebSocket channels.", "Engineered a unified root build script (`npm run build`) compiling both client and server. Express 5 acts as reverse proxy, serving Vite's compiled dist/ directory statically and routing API/WS paths cleanly.", c_teal),
        ("2. Supabase SSL Pooler Rejection", "The Node.js pg driver crashed with SELF_SIGNED_CERT_IN_CHAIN when connecting to the Supabase transaction pooler.", "Implemented dynamic SSL configuration in server/src/db.ts detecting DATABASE_URL and enforcing { rejectUnauthorized: false } for seamless cloud pooler connectivity.", c_blue),
        ("3. Python 3.12 AI C-Library Wheels", "InspireFace native C-bindings lack prebuilt binaries for Python 3.13, causing segmentation faults on newer runtimes.", "Configured dedicated Python 3.12 environment and introduced a thread synchronization lock (_session_lock) around inference calls to guarantee thread safety.", c_emerald),
        ("4. Camera Stream Lifecycle Leaks", "Background tabs or switching views left hardware webcam tracks active, wasting battery and triggering security indicators.", "Implemented HTML5 visibilitychange listeners and ended event handlers that instantly terminate all media stream tracks when the user changes tabs or closes the card.", c_amber),
        ("5. Steganographic Visual Artifacts", "Traditional sequential LSB insertion produces noticeable rectangular bands or color shifts in image gradients.", "Developed a randomized NumPy PRNG bit distribution seeded by a 16-byte cryptographic token, scattering bits across RGB channels with zero perceptible change.", c_teal),
    ]

    for i, (c_name, c_prob, c_sol, c_col) in enumerate(challenges):
        cy = Inches(1.4) + i * Inches(1.08)
        add_card(s12, Inches(0.8), cy, Inches(11.733), Inches(0.98), title=c_name, border_color=c_col)
        
        tb = s12.shapes.add_textbox(Inches(1.0), cy + Inches(0.32), Inches(11.333), Inches(0.6))
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
        r2.font.color.rgb = c_slate400

        r3 = p.add_run()
        r3.text = "→ Solution: "
        r3.font.bold = True
        r3.font.size = Pt(8)
        r3.font.color.rgb = c_emerald

        r4 = p.add_run()
        r4.text = c_sol
        r4.font.size = Pt(8)
        r4.font.color.rgb = c_white

    # =========================================================================
    # SLIDE 13: RESULTS, FUTURE ENHANCEMENTS & ROADMAP
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    add_background(s13)
    add_top_bar(s13, "Results, Outcomes, Limitations & Future Roadmap")

    # Left Card: Results & Outcomes
    add_card(s13, Inches(0.8), Inches(1.4), Inches(5.75), Inches(5.4), title="ACHIEVEMENTS & MEASURED RESULTS", border_color=c_emerald)
    tb_res = s13.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.35), Inches(4.8))
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
        r1.font.size = Pt(9)
        r1.font.color.rgb = c_emerald

        r2 = p.add_run()
        r2.text = r_desc
        r2.font.size = Pt(8)
        r2.font.color.rgb = c_slate200

    # Right Card: Limitations & Future Roadmap
    add_card(s13, Inches(6.78), Inches(1.4), Inches(5.75), Inches(5.4), title="LIMITATIONS & FUTURE ROADMAP", border_color=c_blue)
    tb_fut = s13.shapes.add_textbox(Inches(6.98), Inches(1.85), Inches(5.35), Inches(4.8))
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
        r1.font.size = Pt(9)
        r1.font.color.rgb = c_blue

        r2 = p.add_run()
        r2.text = m_desc
        r2.font.size = Pt(8)
        r2.font.color.rgb = c_slate200

    # =========================================================================
    # SLIDE 14: CONCLUSION & LIVE INTERACTIVE DEMO LINKS
    # =========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    add_background(s14)
    add_top_bar(s14, "Conclusion & Interactive Live Demonstration")

    # Conclusion Card
    c_card_box = add_card(s14, Inches(0.8), Inches(1.4), Inches(11.733), Inches(2.3), title="FINAL PROJECT CONCLUSION")
    tb_c = s14.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(11.333), Inches(1.7))
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
    r_c.font.color.rgb = c_slate200

    # Interactive Action Buttons Row
    links = [
        ("🚀 Launch Live Cloud Application", "Public Render Deployment", "https://chatvia-app-k7o7.onrender.com/", c_teal),
        ("📂 GitHub Source Code Repository", "Full Source & History", "https://github.com/Kishan-Kumar-Hv/hmit", c_blue),
        ("📄 Complete Project Documentation", "Phase 4 Master PDF Report", "Phase_4_Final_Project_Submission_Veltrox.pdf", c_emerald),
    ]

    for i, (l_title, l_sub, l_url, l_col) in enumerate(links):
        lx = Inches(0.8) + i * Inches(3.98)
        l_btn = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, lx, Inches(3.95), Inches(3.78), Inches(1.3))
        l_btn.fill.solid()
        l_btn.fill.fore_color.rgb = c_card
        l_btn.line.color.rgb = l_col
        l_btn.line.width = Pt(1.5)

        tf_l = l_btn.text_frame
        tf_l.vertical_anchor = MSO_ANCHOR.MIDDLE
        p1 = tf_l.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        r1 = p1.add_run()
        r1.text = l_title + "\n"
        r1.font.bold = True
        r1.font.size = Pt(11)
        r1.font.color.rgb = c_white
        r1.hyperlink.address = l_url

        p2 = tf_l.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = l_sub
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = l_col

    # Team Sign-off Card
    s_box = add_card(s14, Inches(0.8), Inches(5.45), Inches(11.733), Inches(1.35), border_color=c_teal)
    tb_s = s14.shapes.add_textbox(Inches(1.0), Inches(5.55), Inches(11.333), Inches(1.15))
    tb_s.text_frame.word_wrap = True
    p_s = tb_s.text_frame.paragraphs[0]
    r_s1 = p_s.add_run()
    r_s1.text = "SUBMITTED BY: TEAM 65 — VELTROX\n"
    r_s1.font.bold = True
    r_s1.font.size = Pt(11)
    r_s1.font.color.rgb = c_teal

    r_s2 = p_s.add_run()
    r_s2.text = (
        "• Kishan Kumar H V (Lead & Cloud/DevOps Architect)    "
        "• Santhosh J R (Full-Stack & Biometrics Core Engineer)\n"
        "• Suhas B R (Security & Steganography Architect)       "
        "• Dhruva Kumar V M (Frontend Lead & UI/UX Specialist)\n"
        "HackMITTen 2026 | Cybersecurity & Applied AI Track | Final Phase 4 Submission Verified"
    )
    r_s2.font.size = Pt(9)
    r_s2.font.color.rgb = c_slate200

    prs.save(filename)
    print(f"Successfully generated presentation deck: {filename}")

if __name__ == "__main__":
    output_file = "Phase_4_Final_Project_Presentation_Team65_Veltrox.pptx"
    if len(sys.argv) > 1:
        output_file = sys.argv[1]
    create_deck(output_file)
