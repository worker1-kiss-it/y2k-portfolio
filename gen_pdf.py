import os
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.colors import HexColor, white, Color
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PIL import Image

BASE = r"C:\git\y2k-portfolio"
OUT = os.path.join(BASE, "Y2K-AI-Portfolio.pdf")
LOGO_W = os.path.join(BASE, "y2klogo.png")       # white
LOGO_B = os.path.join(BASE, "y2klogoblack.png")  # black
BANNER = os.path.join(BASE, "bg1528x1080.png")
W, H = landscape(A4)  # 842 x 595

# Refined palette
DARK = HexColor("#0b1120")
PANEL = HexColor("#131d33")
PANEL2 = HexColor("#0f1729")
ACCENT = HexColor("#00c8f0")
ACCENT_DIM = HexColor("#0090a8")
TEXT = HexColor("#e8e8e8")
DIM = HexColor("#aaaaaa")
MUTED = HexColor("#777777")
FAINT = HexColor("#3a3a3a")
SUBTLE = HexColor("#555555")

def bg(c):
    c.setFillColor(DARK)
    c.rect(0, 0, W, H, fill=1, stroke=0)

def banner_bg(c, alpha=0.78):
    try:
        img = Image.open(BANNER)
        iw, ih = img.size
        scale = max(W/iw, H/ih)
        dw, dh = iw*scale, ih*scale
        dx, dy = (W-dw)/2, (H-dh)/2
        c.drawImage(ImageReader(BANNER), dx, dy, dw, dh)
        c.setFillColor(Color(11/255, 17/255, 32/255, alpha=alpha))
        c.rect(0, 0, W, H, fill=1, stroke=0)
    except: bg(c)

def hl(c, y, x1=60, x2=None, col=None):
    c.setStrokeColor(col or ACCENT)
    c.setLineWidth(1.5)
    c.line(x1, y, x2 or W-60, y)

def tl(c, y, x1=60, x2=None):
    c.setStrokeColor(FAINT)
    c.setLineWidth(0.5)
    c.line(x1, y, x2 or W-60, y)

def t(c, x, y, s, sz=11, col=TEXT, b=False, i=False):
    fn = "Helvetica"
    if b and i: fn = "Helvetica-BoldOblique"
    elif b: fn = "Helvetica-Bold"
    elif i: fn = "Helvetica-Oblique"
    c.setFont(fn, sz)
    c.setFillColor(col)
    c.drawString(x, y, s)

def ct(c, y, s, sz=11, col=TEXT, b=False):
    c.setFont("Helvetica-Bold" if b else "Helvetica", sz)
    c.setFillColor(col)
    c.drawCentredString(W/2, y, s)

def w(c, x, y, text, sz=10.5, col=TEXT, mw=680, ld=15, b=False, i=False):
    fn = "Helvetica"
    if b and i: fn = "Helvetica-BoldOblique"
    elif b: fn = "Helvetica-Bold"
    elif i: fn = "Helvetica-Oblique"
    c.setFont(fn, sz)
    c.setFillColor(col)
    words = text.split()
    line = ""
    for wd in words:
        test = line + (" " if line else "") + wd
        if c.stringWidth(test, fn, sz) > mw:
            c.drawString(x, y, line)
            y -= ld
            line = wd
        else:
            line = test
    if line:
        c.drawString(x, y, line)
        y -= ld
    return y

def pnl(c, x, y, pw, ph, col=PANEL):
    c.setFillColor(col)
    c.roundRect(x, y, pw, ph, 5, fill=1, stroke=0)

def ftr(c, n=None):
    tl(c, 32)
    c.setFont("Helvetica", 7.5)
    c.setFillColor(MUTED)
    c.drawString(60, 19, "Y2K Global OÜ  ·  www.y2k.global")
    if n: c.drawRightString(W-60, 19, str(n))

def draw_logo(c, path, x, y, w2, h2):
    try:
        logo = ImageReader(path)
        c.drawImage(logo, x, y, w2, h2, preserveAspectRatio=True, mask="auto")
    except: pass

# ==================== PAGES ====================

def cover(c):
    banner_bg(c, 0.72)
    # Accent line at top
    c.setFillColor(ACCENT)
    c.rect(0, H-3, W, 3, fill=1, stroke=0)
    # Logo centered
    draw_logo(c, LOGO_W, W/2-60, H/2+60, 120, 60)
    # Company name
    cy = H/2 + 40
    c.setFont("Helvetica-Bold", 32)
    c.setFillColor(white)
    c.drawCentredString(W/2, cy, "Y2K Global")
    cy -= 28
    hl(c, cy, W/2-90, W/2+90)
    cy -= 24
    c.setFont("Helvetica", 15)
    c.setFillColor(ACCENT)
    c.drawCentredString(W/2, cy, "AI Portfolio")
    cy -= 24
    ct(c, cy, "AI Solutions & Consulting", 11, DIM)
    # Bottom block
    by = 80
    ct(c, by+20, "Gergo Kiss, MSc  &  Ing. Ramazan Yildirim", 11, TEXT)
    ct(c, by, "+43 1 442 20 143  ·  office@y2k.global  ·  www.y2k.global", 9, MUTED)
    c.showPage()

def about(c):
    bg(c)
    # Accent bar top
    c.setFillColor(ACCENT)
    c.rect(0, H-3, W, 3, fill=1, stroke=0)
    t(c, 60, H-50, "About Y2K Global", 22, ACCENT, True)
    hl(c, H-60)
    y = H-82
    y = w(c, 60, y, "Y2K Global is a technology consultancy that turns complex business challenges into intelligent, automated systems. We specialize in artificial intelligence, machine learning, and software optimization — delivering solutions that create measurable value, not just impressive demos.", 11, TEXT, W-120, 16)
    y -= 4
    y = w(c, 60, y, "With deep expertise spanning natural language processing, predictive analytics, document intelligence, and full-stack development, we bridge the gap between cutting-edge AI research and production-ready business solutions.", 11, TEXT, W-120, 16)
    y -= 18
    t(c, 60, y, "Leadership", 15, ACCENT, True)
    tl(c, y-8)
    y -= 26
    pnl(c, 60, y-50, W-120, 56)
    t(c, 72, y-6, "Gergo Kiss, MSc  &  Ing. Ramazan Yildirim", 12, TEXT, True)
    t(c, 72, y-20, "Co-Founders", 9.5, ACCENT)
    w(c, 72, y-34, "Gergo and Ramazan co-founded and jointly lead Y2K Global. As former fellow students during their Computer Science studies in Vienna, they developed a shared passion for applying artificial intelligence and machine learning to real-world business challenges. This shared interest ultimately led them to establish a consultancy that combines academic insight with practical engineering.", 9, DIM, W-155, 12)
    y -= 68
    t(c, 60, y, "How We Work", 15, ACCENT, True)
    tl(c, y-8)
    y -= 28
    phases = [
        ("Discover", "Deep immersion in your business context, data, and constraints. We identify where AI creates real value."),
        ("Architect", "Systems that integrate with existing infrastructure, meet security requirements, and scale with your growth."),
        ("Build & Validate", "Iterative development against real data. Every sprint produces working software you can evaluate."),
        ("Deploy & Evolve", "Production deployment with monitoring and ongoing optimization as your needs and data patterns change."),
    ]
    for idx, (title, desc) in enumerate(phases):
        pnl(c, 60, y-26, W-120, 32)
        # Number circle
        cx2 = 78
        cy2 = y-10
        c.setFillColor(ACCENT)
        c.circle(cx2, cy2, 10, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 9)
        c.setFillColor(DARK)
        c.drawCentredString(cx2, cy2-3, str(idx+1))
        t(c, 96, y-4, title, 10.5, TEXT, True)
        w(c, 96, y-17, desc, 9, DIM, W-175, 12)
        y -= 38
    ftr(c, 2)
    c.showPage()

def cs(c, num, total, tag, title, challenge, solution, results, techs, pn):
    bg(c)
    # Accent bar top
    c.setFillColor(ACCENT)
    c.rect(0, H-3, W, 3, fill=1, stroke=0)
    # Header
    t(c, 60, H-38, f"Case Study {num:02d} / {total:02d}", 9, MUTED)
    t(c, 60, H-60, title, 20, TEXT, True)
    t(c, 60, H-78, tag, 10, ACCENT, i=True)
    hl(c, H-86)
    # Two columns
    lw = (W-150) * 0.55
    rw = (W-150) * 0.42
    rx = 90 + lw
    # Left: challenge + solution
    y = H-106
    t(c, 60, y, "Challenge", 11, ACCENT, True)
    y -= 16
    y = w(c, 60, y, challenge, 10, DIM, lw, 14)
    y -= 14
    t(c, 60, y, "Solution", 11, ACCENT, True)
    y -= 16
    y = w(c, 60, y, solution, 10, DIM, lw, 14)
    # Right: outcomes + tech
    ry = H-106
    t(c, rx, ry, "Outcomes", 11, ACCENT, True)
    ry -= 18
    for r2 in results:
        t(c, rx+2, ry, "→", 9, ACCENT_DIM)
        ry = w(c, rx+16, ry, r2, 10, TEXT, rw-16, 14)
        ry -= 5
    ry -= 14
    t(c, rx, ry, "Technologies", 11, ACCENT, True)
    ry -= 18
    tx2 = rx
    for te in techs:
        tw2 = c.stringWidth(te, "Helvetica", 8.5) + 14
        if tx2 + tw2 > W-60:
            tx2 = rx
            ry -= 20
        pnl(c, tx2, ry-5, tw2, 17)
        t(c, tx2+7, ry-1, te, 8.5, DIM)
        tx2 += tw2 + 6
    ftr(c, pn)
    c.showPage()

def tech_stack(c, pn):
    bg(c)
    c.setFillColor(ACCENT)
    c.rect(0, H-3, W, 3, fill=1, stroke=0)
    t(c, 60, H-50, "Technology Stack", 22, ACCENT, True)
    hl(c, H-60)
    y = H-88
    stacks = [
        ("AI & Machine Learning", "Python  ·  TensorFlow  ·  PyTorch  ·  scikit-learn  ·  Hugging Face  ·  LangChain  ·  OpenAI  ·  Claude"),
        ("Natural Language Processing", "spaCy  ·  NLTK  ·  Transformers  ·  RAG Pipelines  ·  ChromaDB  ·  Pinecone  ·  Semantic Search"),
        ("Backend Development", "FastAPI  ·  Django  ·  ASP.NET Core  ·  Entity Framework  ·  Node.js  ·  Express"),
        ("Frontend & Visualization", "React  ·  Next.js  ·  Vue.js  ·  Nuxt  ·  Tailwind CSS  ·  D3.js  ·  Plotly"),
        ("Databases", "PostgreSQL  ·  SQL Server  ·  MariaDB  ·  Redis  ·  MongoDB  ·  Elasticsearch"),
        ("Infrastructure", "Docker  ·  Linux  ·  Nginx  ·  CI/CD  ·  Azure DevOps  ·  GitHub Actions"),
        ("Document Processing", "Tesseract OCR  ·  pdf2image  ·  Apache Tika  ·  Custom Pipelines"),
        ("Integration", "REST  ·  GraphQL  ·  WebSocket  ·  IMAP/SMTP  ·  OAuth 2.0  ·  SSO"),
    ]
    for title, items in stacks:
        pnl(c, 60, y-22, W-120, 30)
        t(c, 72, y-2, title, 10.5, ACCENT, True)
        t(c, 230, y-2, items, 9, DIM)
        y -= 40
    # Small logo at bottom
    draw_logo(c, LOGO_W, W/2-25, 50, 50, 25)
    ftr(c, pn)
    c.showPage()

def contact(c, pn):
    banner_bg(c, 0.78)
    c.setFillColor(ACCENT)
    c.rect(0, H-3, W, 3, fill=1, stroke=0)
    # White logo top
    draw_logo(c, LOGO_W, W/2-50, H-120, 100, 50)
    # CTA
    cy = H/2 + 60
    ct(c, cy, "Let’s Build Something Intelligent Together", 22, ACCENT, True)
    cy -= 14
    hl(c, cy, W/2-100, W/2+100)
    cy -= 24
    ct(c, cy, "Get in touch for a free initial consultation.", 12, TEXT)
    # Founders side by side
    sy = H/2 - 30
    lx = 160
    rx = W/2 + 40
    # Gergo
    t(c, lx, sy+6, "Gergo Kiss, MSc", 13, TEXT, True)
    t(c, lx, sy-8, "Managing Director · AI Systems Architect", 10, DIM)
    t(c, lx, sy-22, "g.kiss@y2k.global", 9, MUTED)
    # Rami
    t(c, rx, sy+6, "Ing. Ramazan Yildirim", 13, TEXT, True)
    t(c, rx, sy-8, "Managing Director · AI Solutions Engineer", 10, DIM)
    t(c, rx, sy-22, "r.yildirim@y2k.global", 9, MUTED)
    sy -= 36
    ct(c, sy, "+43 1 442 20 143  ·  office@y2k.global  ·  www.y2k.global", 9, MUTED)
    sy -= 44
    tl(c, sy, 250, W-250)
    sy -= 18
    lx = 250
    rx2 = W/2 + 10
    t(c, rx2, sy, "In cooperation with", 9, MUTED)
    sy -= 14
    t(c, lx, sy, "Y2K Global OÜ", 10, DIM, True)
    t(c, rx2, sy, "KISS IT Solutions e.U.", 10, DIM, True)
    sy -= 14
    t(c, lx, sy, "Sepapaja 6 · 15551 Tallinn · Estonia", 9, MUTED)
    t(c, rx2, sy, "Seitenstettengasse 5/37 · 1010 Vienna · Austria", 9, MUTED)
    sy -= 14
    t(c, lx, sy, "Reg. No.: 17373395 · VAT: EE102928187", 9, MUTED)
    t(c, rx2, sy, "Reg. No.: FN606200x · VAT: ATU68895724", 9, MUTED)
    ftr(c, pn)
    c.showPage()

def main():
    c = canvas.Canvas(OUT, pagesize=landscape(A4))
    c.setTitle("Y2K Global - AI Portfolio")
    c.setAuthor("Y2K Global")
    c.setSubject("AI Solutions & Consulting")
    cover(c)
    about(c)
    cases = [
        ("Funding & Grants Automation",
         "AI-Powered Funding Discovery & Application Platform",
         "European organizations spend hundreds of hours quarterly searching for relevant funding opportunities across fragmented EU and national databases. Applications are prepared manually, deadlines are missed, and success rates suffer from poor opportunity-fit matching.",
         "We developed an end-to-end AI platform that continuously monitors EU portals, national databases, and email communications for funding opportunities. The system uses NLP to evaluate each opportunity against the organization’s profile, scores relevance and feasibility, generates comprehensive application briefs, and manages the full lifecycle from discovery through submission and reporting.",
         ["Automated scanning of 100+ funding sources", "AI-driven opportunity matching and scoring", "Intelligent application brief generation", "Full lifecycle tracking and management", "Significant reduction in manual research hours"],
         ["Python", "FastAPI", "PostgreSQL", "NLP", "Claude AI", "Email/IMAP", "Web Scraping"]),
        ("Private AI Infrastructure",
         "Private AI Platform with Standard SDK Integration",
         "Organizations needed to deploy AI and large language model capabilities within their own infrastructure — without sending sensitive data to external cloud providers. Existing solutions required deep AI expertise and custom integration for each use case, creating a barrier to adoption across departments.",
         "We built a private AI platform that runs entirely within the client’s infrastructure, providing an OpenAI-compatible API as a standard SDK interface. This means existing tools and applications can connect with zero code changes. The platform features intelligent multi-model routing for optimal cost and performance, a comprehensive admin dashboard for usage monitoring, and production-grade access controls. Teams across the organization can leverage AI capabilities through familiar interfaces while keeping all data in-house.",
         ["Fully private deployment — no data leaves the organization", "OpenAI-compatible SDK for seamless integration", "Multi-model routing for cost/performance optimization", "Admin dashboard for monitoring and governance", "Zero-friction adoption across departments"],
         ["Python", "LLM Inference", "API Gateway", "Docker", "REST API", "Admin UI"]),
        ("System Optimization & AI",
         "Fleet Management Platform — Performance & AI Enhancement",
         "A fleet management platform serving hundreds of locations and multiple operators suffered from critical performance bottlenecks. Key operational reports took over 14 seconds to load. Data type mismatches in the engineering pipeline caused intermittent system failures that eroded user trust.",
         "We performed a deep technical audit of the .NET/SQL Server stack, identified root causes in query design, database indexing, and ORM configuration. Applied surgical optimizations: replaced expensive SQL functions, added covering indexes, implemented NoTracking for read operations, and refreshed stale view metadata. Then extended the system with an AI companion featuring natural language fleet data querying and vision-based document processing with a ReAct agent architecture.",
         ["87% faster report generation (14.7s → 1.9s)", "Critical data pipeline errors resolved", "Natural language querying of fleet data", "Vision-based document processing", "ReAct agent with safety guardrails"],
         [".NET", "Entity Framework Core", "SQL Server", "Python", "Vision AI", "NLP"]),
        ("Financial Process Automation",
         "AI-Powered Financial Automation Platform",
         "Financial service providers and bookkeepers managing multiple clients were drowning in repetitive manual work: invoice data entry, transaction categorization, and report generation. Each client had unique requirements, making standardized automation difficult. The initial challenge emerged during academic research into applying machine learning to accounting workflows.",
         "What began as a master’s thesis project exploring ML applications in accounting automation evolved into a comprehensive financial automation platform. The system combines AI-powered invoice extraction using custom OCR pipelines, automated transaction categorization via purpose-built ML classification models, anomaly detection for quality assurance, and intelligent financial reporting. Built as a multi-tenant SaaS platform, each client has fully isolated data with configurable workflows, chart-of-accounts mappings, and export capabilities including Austrian UVA tax statements.",
         ["AI-powered invoice extraction (OCR + NLP)", "ML-based transaction categorization", "Automated anomaly detection", "Multi-tenant architecture with data isolation", "Configurable workflows per client", "Academic research foundation with industry validation"],
         ["Python", "FastAPI", "Next.js", "PostgreSQL", "scikit-learn", "Tesseract OCR", "Redis"]),
        ("SME Business Automation",
         "AI-Assisted Offer & Quote Creation for SMEs",
         "Small and medium-sized enterprises spend disproportionate time creating offers and quotes. Each proposal requires pulling together service descriptions, calculating pricing, applying client-specific terms, and formatting professionally — a tedious process that takes hours and is prone to inconsistency, especially for businesses without dedicated sales teams.",
         "We developed an AI-assisted offer creation tool designed specifically for SMEs. The system guides users through a structured workflow, suggests relevant service descriptions and pricing based on historical patterns, auto-populates client data, and generates professional, branded PDF proposals. AI assists at every step — from recommending line items to optimizing pricing based on past acceptance rates — while keeping the user in full control of the final output.",
         ["Guided offer creation workflow for non-technical users", "AI-powered pricing and line item suggestions", "Professional branded PDF generation", "Client history and pattern learning", "Designed for SMEs without dedicated sales teams"],
         ["Python", "FastAPI", "React", "NLP", "PDF Generation", "ML"]),
        ("Data Analytics & Monitoring",
         "Operational Intelligence & Predictive Monitoring Platform",
         "Organizations managing distributed infrastructure — from fleet operations to facility networks — lacked real-time visibility into operational performance. Data from sensors, devices, and operational systems sat in silos, making it impossible to detect emerging issues before they became costly failures. Decision-making relied on periodic manual reporting rather than live intelligence.",
         "We built a data analytics and monitoring platform that ingests operational data from multiple sources in real time, applies statistical analysis and machine learning models for anomaly detection and trend forecasting, and presents actionable insights through interactive dashboards. The platform supports configurable alerting thresholds, predictive maintenance scheduling, and automated reporting. Currently deployed at a smaller scale for infrastructure monitoring, the architecture is designed to scale to large multi-site operations.",
         ["Real-time operational data ingestion and analysis", "ML-based anomaly detection and alerting", "Predictive maintenance and trend forecasting", "Interactive dashboards with drill-down capability", "Scalable architecture for multi-site deployment"],
         ["Python", "Time-Series DB", "ML", "FastAPI", "React", "Docker", "IoT Integration"]),
        ("AI Agent Orchestration",
         "AI Agent Deployment & Orchestration Platform",
         "Organizations adopting AI assistants face a scaling challenge: each agent instance requires manual setup, credential management, model configuration, and infrastructure provisioning. Managing dozens of AI agents across teams becomes an operational burden, with no unified visibility into agent status, resource usage, or inter-agent coordination.",
         "We built a web-based mission control platform for deploying, managing, and orchestrating multiple AI agent instances as isolated Docker containers. The system provides one-click deployment with automatic configuration generation, a secure credential vault with AES-256 encryption, intelligent reverse proxy with WebSocket bridging for seamless dashboard access, and real-time status monitoring across all agents. The architecture supports IP whitelisting, TLS termination, and capacity planning for scaling to dozens of concurrent agent instances on a single host.",
         ["One-click agent deployment with full isolation", "Secure credential vault and model configuration", "Real-time mission control dashboard", "Intelligent reverse proxy with WebSocket MITM", "Scalable to 25+ concurrent agents per host"],
         ["TypeScript", "Node.js", "Express", "React", "Docker", "SQLite", "WebSocket", "TLS"]),
    ]
    total = len(cases)
    for idx, (tag, title, chal, sol, res, tech) in enumerate(cases):
        cs(c, idx+1, total, tag, title, chal, sol, res, tech, idx+3)
    tech_stack(c, total+3)
    contact(c, total+4)
    c.save()
    pages = total + 4
    sz = os.path.getsize(OUT)
    print(f"PDF generated: {OUT}")
    print(f"Pages: {pages}")
    print(f"Size: {sz:,} bytes ({sz/1024:.1f} KB)")

if __name__ == "__main__":
    main()
