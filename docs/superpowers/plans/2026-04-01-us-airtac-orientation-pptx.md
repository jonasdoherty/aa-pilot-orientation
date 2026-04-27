# US Air Attack Pilot Orientation PPTX — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a 30-slide PowerPoint presentation for Coulson Aviation's US air attack pilot orientation, adapted from the existing Australian ground school deck.

**Architecture:** A Python script (`build_presentation.py`) uses the existing `Air Attack Pilot Ground School.pptx` as a template to inherit the "Coulson Carbon" theme, Garamond fonts, and dark parallelogram master slide design. The script clears all existing slides, then builds 30 new slides using Layout 0 (Title Slide) for the title and section transitions, and Layout 1 (Title and Content) for all content slides. Speaker notes are added to every slide.

**Tech Stack:** Python 3, python-pptx 1.0.2 (already installed)

**Spec:** `docs/superpowers/specs/2026-04-01-us-airtac-pilot-orientation-design.md`

---

### Template Details (reference for all tasks)

- **Source template:** `Air Attack Pilot Ground School.pptx`
- **Output file:** `US Air Attack Pilot Orientation.pptx`
- **Slide size:** 13.33" x 7.50" (16:9 widescreen)
- **Theme:** "Coulson Carbon" — Garamond font, dark industrial background with angled parallelogram bands
- **Layout 0** ("Title Slide"): CENTER_TITLE placeholder (idx 0), SUBTITLE placeholder (idx 1)
- **Layout 1** ("Title and Content"): TITLE placeholder (idx 0), OBJECT/content placeholder (idx 1)
- **Coulson logo:** On slide 1 of the template — Picture 3 at position (6566088, 1271851), size (4626568, 3575075), duotone + 85% alpha. Extract and reuse on new title slide.

---

### Task 1: Script setup and template handling

**Files:**
- Create: `build_presentation.py`

- [ ] **Step 1: Create the script with template loading and slide clearing**

```python
"""
Build the US Air Attack Pilot Orientation PowerPoint.
Uses the existing Australian ground school deck as a template for theme/branding.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN
import copy
import os

TEMPLATE = r"Air Attack Pilot Ground School.pptx"
OUTPUT = r"US Air Attack Pilot Orientation.pptx"

def load_template(path):
    """Load the template PPTX and extract the Coulson logo image from slide 1."""
    prs = Presentation(path)
    
    # Extract logo image bytes from slide 1 (Picture 3)
    logo_blob = None
    logo_image_part = None
    for shape in prs.slides[0].shapes:
        if shape.name == "Picture 3" and shape.shape_type == 13:
            logo_blob = shape.image.blob
            logo_image_part = shape.image.content_type
            break
    
    # Delete all existing slides
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        prs.part.drop_rel(rId)
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])
    
    return prs, logo_blob, logo_image_part

def add_content_slide(prs, title_text, bullets, notes_text=None):
    """Add a Title and Content slide (Layout 1) with bullet points and optional speaker notes."""
    layout = prs.slide_layouts[1]  # "Title and Content"
    slide = prs.slides.add_slide(layout)
    
    # Set title
    slide.shapes.title.text = title_text
    
    # Set bullet content
    tf = slide.placeholders[1].text_frame
    tf.clear()
    for i, bullet in enumerate(bullets):
        if i == 0:
            para = tf.paragraphs[0]
        else:
            para = tf.add_paragraph()
        para.text = bullet["text"]
        para.level = bullet.get("level", 0)
    
    # Add speaker notes
    if notes_text:
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = notes_text
    
    return slide

def add_title_slide(prs, title_text, subtitle_text=None, notes_text=None):
    """Add a Title Slide (Layout 0) for title page or section dividers."""
    layout = prs.slide_layouts[0]  # "Title Slide"
    slide = prs.slides.add_slide(layout)
    
    slide.shapes.title.text = title_text
    if subtitle_text and 1 in slide.placeholders:
        slide.placeholders[1].text = subtitle_text
    
    if notes_text:
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = notes_text
    
    return slide

# --- Build the presentation ---
prs, logo_blob, logo_content_type = load_template(TEMPLATE)
```

- [ ] **Step 2: Run the script to verify template loading works**

Run: `cd "/c/Users/jonas/Projects/AA Pilot Training" && python3 build_presentation.py`
Expected: No errors. No output file yet (we haven't saved).

---

### Task 2: Section 1 — Opening (Slides 1-3)

**Files:**
- Modify: `build_presentation.py`

- [ ] **Step 1: Add the opening section slides**

Append to `build_presentation.py` before any save logic:

```python
# ============================================================
# SECTION 1: OPENING
# ============================================================

# Slide 1 — Title
slide1 = add_title_slide(
    prs,
    title_text="Air Attack Pilot Orientation",
    subtitle_text="Coulson Aviation — US Operations",
    notes_text=(
        "Welcome everyone. This session is about getting aligned before you go out "
        "and start flying. You've been through the technical ground school — this is "
        "about how Coulson does things, what we expect from you, and how to make the "
        "most of the air attack role."
    )
)

# Add Coulson logo to title slide
if logo_blob:
    import io
    slide1.shapes.add_picture(
        io.BytesIO(logo_blob),
        left=Emu(6566088), top=Emu(1271851),
        width=Emu(4626568), height=Emu(3575075)
    )

# Slide 2 — Why We're Here
add_content_slide(prs,
    title_text="Why We're Here",
    bullets=[
        {"text": "You're about to go out on your own — single pilot, representing Coulson"},
        {"text": "This session is about getting everyone aligned before you do"},
        {"text": "No captain to shadow — what you learn in this room is what you take to the field"},
    ],
    notes_text=(
        "Unlike our tanker crews, where a new pilot flies with an experienced captain, "
        "you're going out solo. There's no one in the left seat to show you the ropes. "
        "So we need to make sure you leave this room knowing how we operate, what we "
        "expect, and who to call when you need help."
    )
)

# Slide 3 — The Opportunity
add_content_slide(prs,
    title_text="The Opportunity",
    bullets=[
        {"text": "The air attack seat gives you something no other role does:"},
        {"text": "4-hour missions with an experienced ATGS beside you", "level": 1},
        {"text": "Watching fire behavior develop over hours, not minutes", "level": 1},
        {"text": "Hearing every tactical decision, seeing how retardant works on the fire", "level": 1},
        {"text": "Once you're in the tanker: 5-minute snapshots, no ATGS beside you, no time to watch things develop"},
        {"text": "After a couple of seasons, you should be able to teach the class on aerial firefighting tactics and fire behavior"},
        {"text": "Career pathway: merit-based progression to air tanker — promote from within, merit first, seniority as tiebreaker"},
    ],
    notes_text=(
        "This is the message I really want to land. The air attack seat is the best "
        "seat in the house for learning aerial firefighting. You're overhead for four "
        "hours at a time. You've got one of the most experienced ground firefighters "
        "in the country sitting next to you. You're watching fire behavior develop. "
        "You're hearing target descriptions. You're seeing how tactical decisions play "
        "out. When you move to the tanker, that's all gone — you're on scene for five "
        "minutes and then you're back to the reload base. Use this time. It's more "
        "valuable than most people realize until it's gone."
    )
)
```

- [ ] **Step 2: Add save logic at the bottom of the script**

Append to the very end of `build_presentation.py`:

```python
# ============================================================
# SAVE
# ============================================================
prs.save(OUTPUT)
print(f"Saved: {OUTPUT} ({len(prs.slides)} slides)")
```

- [ ] **Step 3: Run and verify**

Run: `cd "/c/Users/jonas/Projects/AA Pilot Training" && python3 build_presentation.py`
Expected: `Saved: US Air Attack Pilot Orientation.pptx (3 slides)`

Open the file to verify: title slide with logo, "Why We're Here" slide, "The Opportunity" slide. Coulson Carbon theme should be visible.

---

### Task 3: Section 2 — The Coulson Standard (Slides 4-8)

**Files:**
- Modify: `build_presentation.py` (insert before SAVE section)

- [ ] **Step 1: Add the Coulson Standard slides**

```python
# ============================================================
# SECTION 2: THE COULSON STANDARD
# ============================================================

# Slide 4 — What Coulson Expects
add_content_slide(prs,
    title_text="What Coulson Expects",
    bullets=[
        {"text": "Professionalism, initiative, consistency"},
        {"text": "You are Coulson's face at every tanker base, every agency interaction, every radio call"},
        {"text": "Approach the role with humility and a hunger to learn"},
    ],
    notes_text=(
        "You represent this company every time you walk onto a tanker base, every time "
        "you key the mic, every interaction you have with agency personnel. How you "
        "carry yourself matters. We've built a strong reputation and we need you to "
        "uphold it."
    )
)

# Slide 5 — Tanker Base Conduct
add_content_slide(prs,
    title_text="Tanker Base Conduct",
    bullets=[
        {"text": "You are a guest at these bases"},
        {"text": "Dress code: Coulson polo, closed-toe shoes — present yourself well"},
        {"text": "Clean up after yourself, be polite, represent the company"},
        {"text": "Build good relationships with techs, tanker base personnel, agency staff"},
    ],
    notes_text=(
        "We are guests at every tanker base we operate from. Clean up after yourself. "
        "Be polite. Wear the Coulson polo and closed-toe shoes — present yourself in a "
        "way that reflects well on the company. Build relationships with the people you "
        "work around. These are the techs who keep your airplane flying, the base "
        "personnel who support your operation, and the agency staff who make it all "
        "happen. Treat them well."
    )
)

# Slide 6 — Your Support Structure
add_content_slide(prs,
    title_text="Your Support Structure",
    bullets=[
        {"text": "You're flying alone, but you're not on your own"},
        {"text": "Chief Pilot, Chief of Training, Assistant Chief Pilot — open door"},
        {"text": "We hold you accountable AND we have your back"},
        {"text": "Share struggles early — we want to hear from you before things become problems"},
    ],
    notes_text=(
        "Just because you're single pilot doesn't mean you're on your own. You can "
        "come to me, you can talk to our Chief of Training, you can talk to your "
        "Assistant Chief Pilot. We have an open door. If something is going on in the "
        "field — a difficult ATGS relationship, an operational issue, something you're "
        "unsure about — tell us. We'd rather hear about it early and help you work "
        "through it than find out later when it's become a real problem. We're going "
        "to hold you accountable, but we're also going to have your back."
    )
)

# Slide 7 — Being Involved
add_content_slide(prs,
    title_text="Being Involved",
    bullets=[
        {"text": "Don't just fly the plane — learn the mission"},
        {"text": "Pay attention to:", "level": 0},
        {"text": "Fire behavior development over the course of a mission", "level": 1},
        {"text": "Tactical decisions — why the ATGS puts retardant where they do", "level": 1},
        {"text": "Target descriptions and how they translate to what you see", "level": 1},
        {"text": "How retardant interacts with fire", "level": 1},
        {"text": "You are part of the aerial supervision team — act like it"},
    ],
    notes_text=(
        "This is probably the most important slide in this whole presentation. Your "
        "primary job is to fly the airplane and provide a stable platform. But if "
        "that's ALL you do, you're missing the point. You're overhead for four hours "
        "at a time. You've got a front-row seat to watch fire behavior develop, to "
        "hear how tactical decisions are made, to see how retardant works on the fire. "
        "This is your education for the tanker. The pilots who get the most out of "
        "this role are the ones who are paying attention to all of it — not just "
        "flying the orbit."
    )
)

# Slide 8 — The ATGS Relationship
add_content_slide(prs,
    title_text="The ATGS Relationship",
    bullets=[
        {"text": "Ask an ATGS who their favorite pilot is — the answer is almost always: the one who was involved, hungry to learn"},
        {"text": "Ask questions at the right times: transit, downtime, debrief — not mid-mission"},
        {"text": "Be curious — build the relationship"},
        {"text": "Debrief after every mission — this is where the real learning happens"},
        {"text": "The fire world is small — your reputation follows you"},
    ],
    notes_text=(
        "If you ask an ATGS who their favorite pilot is, they'll almost always tell "
        "you it's the one who was engaged. The one who wanted to learn. The one who "
        "asked questions during transit and at the debrief. They value that, and those "
        "relationships matter in this business. The fire world is small. Your "
        "reputation — good or bad — follows you. Be the pilot people want to work with."
    )
)
```

- [ ] **Step 2: Run and verify**

Run: `cd "/c/Users/jonas/Projects/AA Pilot Training" && python3 build_presentation.py`
Expected: `Saved: US Air Attack Pilot Orientation.pptx (8 slides)`

---

### Task 4: Section 3 — How We Operate (Slides 9-14)

**Files:**
- Modify: `build_presentation.py` (insert before SAVE section)

- [ ] **Step 1: Add the How We Operate slides**

```python
# ============================================================
# SECTION 3: HOW WE OPERATE
# ============================================================

# Slide 9 — Daily Readiness
add_content_slide(prs,
    title_text="Daily Readiness",
    bullets=[
        {"text": "Show up early — aircraft ready to go at start time, not deep in preflight at start time"},
        {"text": "Continuous weather and regional awareness throughout the day"},
        {"text": "Gear staged, ready for immediate dispatch with possible diversion"},
        {"text": "Fuel planning: calculate MTOW using forecast peak temps, know your performance-limiting airports"},
    ],
    notes_text=(
        "When the availability time hits, you need to be ready to go. That means "
        "showing up early enough to have the airplane preflighted, fueled, and ready. "
        "Not standing out on the ramp doing a walk-around when the phone rings. Stay "
        "aware of the weather, know what's going on in your area, and have your gear "
        "staged so that when dispatch calls, you can be airborne quickly."
    )
)

# Slide 10 — Dispatch and Mission Prep
add_content_slide(prs,
    title_text="Dispatch and Mission Prep",
    bullets=[
        {"text": "FRAT completed in SMS:"},
        {"text": "Before daily strategic briefing", "level": 1},
        {"text": "Updated after dispatch and between missions", "level": 1},
        {"text": "Finalized at end of day", "level": 1},
        {"text": "Pre-mission brief with AAS: objectives, frequencies, workload division, tactical pause criteria"},
        {"text": "Dispatch form + FRAT finalized = cleared for dispatch — not before"},
    ],
    notes_text=(
        "The FRAT is a living document throughout the day. You start it during your "
        "walk-around, update it when conditions change or after dispatch, and finalize "
        "it at end of day. Before you launch on a mission, you brief with your AAS — "
        "what are the objectives, what frequencies are we on, how are we dividing the "
        "workload. You are not cleared for dispatch until both the dispatch form and "
        "FRAT are complete."
    )
)

# Slide 11 — Administrative Standards
add_content_slide(prs,
    title_text="Administrative Standards",
    bullets=[
        {"text": "EFL daily reporting: submitted same evening, every evening"},
        {"text": "Company email signature format — standardized across the fleet"},
        {"text": "What you include in the email body matters"},
        {"text": "Flight time logging: your times must match ATGS records to the minute"},
        {"text": "All times in LOCAL, not UTC"},
    ],
    notes_text=(
        "Admin consistency matters because you're all operating independently. We need "
        "everyone doing things the same way. EFL dailies go out the same evening — not "
        "the next morning, not two days later. Use the standard company email signature. "
        "The content in your daily email matters — include the relevant information. "
        "Your flight times need to match the ATGS records to the minute. If there's a "
        "discrepancy, it creates problems. And remember — all times are local, not UTC."
    )
)

# Slide 12 — Flight Following and Comms
add_content_slide(prs,
    title_text="Flight Following and Comms",
    bullets=[
        {"text": "AFF for all USFS flights — establish with initial radio call"},
        {"text": "WhatsApp group for non-agency and repositioning flights:"},
        {"text": "Aircraft, departure, destination, ETA, POB, endurance", "level": 1},
    ],
    notes_text=(
        "For any flight under a USFS agreement, you'll be on AFF — Automated Flight "
        "Following. For non-agency flights — repositioning, ferry, anything not under "
        "an agency dispatch — use the WhatsApp group. Post your aircraft, where you're "
        "going, ETA, passengers, and fuel endurance. Someone needs to know where you "
        "are at all times."
    )
)

# Slide 13 — Crew Change and Handover
add_content_slide(prs,
    title_text="Crew Change and Handover",
    bullets=[
        {"text": "Crew handover form required at every crew change"},
        {"text": "Emailed to inbound crew, Ops Manager, Chief Pilot, Fleet ACP"},
        {"text": "The inbound crew should be able to step into your shoes with no surprises"},
    ],
    notes_text=(
        "When you hand off to the next crew, the handover form needs to be complete "
        "and sent to the right people. Think about what the inbound pilot needs to "
        "know: aircraft status, any discrepancies, what's been going on operationally, "
        "anything unusual. They should be able to read your handover and step into "
        "the operation seamlessly."
    )
)

# Slide 14 — The Field Guide
add_content_slide(prs,
    title_text="Citation Operations Field Guide",
    bullets=[
        {"text": "Available as a reference for day-to-day operational procedures"},
        {"text": "Covers: fuel procedures, aircraft out-of-service events, FMC updates, base ops, expenses and travel, key contacts"},
    ],
    notes_text=(
        "The Citation Operations Field Guide covers the day-to-day procedures you'll "
        "need to reference — fuel cards, what to do if the airplane goes out of "
        "service, FMC database updates, expense reports, travel booking, and key "
        "contacts. It's there for you when you need it."
    )
)
```

- [ ] **Step 2: Run and verify**

Run: `cd "/c/Users/jonas/Projects/AA Pilot Training" && python3 build_presentation.py`
Expected: `Saved: US Air Attack Pilot Orientation.pptx (14 slides)`

---

### Task 5: Section 4 — US Operations Context (Slides 15-18)

**Files:**
- Modify: `build_presentation.py` (insert before SAVE section)

- [ ] **Step 1: Add the US Operations Context slides**

```python
# ============================================================
# SECTION 4: US OPERATIONS CONTEXT
# ============================================================

# Slide 15 — The US Fire Season
add_content_slide(prs,
    title_text="The US Fire Season",
    bullets=[
        {"text": "[Presenter: update with current season contract information]"},
        {"text": "Exclusive use vs. call when needed — what that means for you"},
        {"text": "Where you can expect to be based and typical deployment patterns"},
        {"text": "Flights outside contract obligations — repositioning, ferry, maintenance"},
        {"text": "Expect: long days, diversions, dispatch distances 10–300 miles, forward basing near active fires"},
    ],
    notes_text=(
        "UPDATE THIS SLIDE EACH SEASON with current contract information.\n\n"
        "Walk through what contracts we have this year, where they are, whether "
        "they're exclusive use or call when needed, and what that means for the "
        "pilots' day-to-day. Talk about where they can expect to be based and the "
        "kinds of flights they may be asked to do outside their normal contract "
        "tasking — repositioning, ferry flights, maintenance runs. Set expectations: "
        "long days, possible diversions, and being ready to go wherever the work is."
    )
)

# Slide 16 — GACC Structure and Dispatch
add_content_slide(prs,
    title_text="GACC Structure and Dispatch",
    bullets=[
        {"text": "Geographic Area Coordination Centers (GACCs) — the national dispatch system"},
        {"text": "NIFC coordination — how national resources are prioritized and moved"},
        {"text": "The dispatch chain: national → geographic area → local unit"},
        {"text": "How you get tasked and what to expect from the process"},
    ],
    notes_text=(
        "The US wildland fire dispatch system is organized around Geographic Area "
        "Coordination Centers — GACCs. There are several across the country, and they "
        "coordinate resource ordering within their areas. Above them is NIFC — the "
        "National Interagency Fire Center — which handles national-level coordination "
        "and resource movement between geographic areas. When things get busy, "
        "resources get moved around the country. Understanding this system helps you "
        "understand why you might get sent somewhere unexpected."
    )
)

# Slide 17 — Agency Relationships
add_content_slide(prs,
    title_text="Agency Relationships",
    bullets=[
        {"text": "USFS, BLM, CAL FIRE — who you may be working for"},
        {"text": "You're operating under their contracts — understand the relationship"},
        {"text": "Agency personnel are your partners, not your customers"},
    ],
    notes_text=(
        "You may be working under contracts with USFS, BLM, CAL FIRE, or other "
        "agencies. Each has their own way of doing things. The key thing to understand "
        "is that these are partnerships. Agency personnel — your ATGS, dispatch, "
        "tanker base managers — are people you work alongside. Treat them as partners. "
        "Build those relationships. It makes everything work better."
    )
)

# Slide 18 — Regulatory Framework
add_content_slide(prs,
    title_text="Regulatory Framework",
    bullets=[
        {"text": "All flights under USFS agreements operate under FAR Part 135"},
        {"text": "Squawk 1255 for firefighting operations"},
        {"text": "TFRs: dispatch to one incident does not clear you through another incident's TFR"},
        {"text": "NWCG PMS 505 — the US standard for aerial supervision"},
    ],
    notes_text=(
        "All of our flights under USFS agreements operate under Part 135 — no "
        "exceptions. Squawk 1255 for firefighting ops. One thing to be aware of when "
        "it gets busy: just because you've been dispatched to an incident doesn't mean "
        "you can fly through another incident's TFR on the way there. Each TFR is its "
        "own airspace restriction. NWCG PMS 505 is the national standard for aerial "
        "supervision — it defines how all of this is supposed to work."
    )
)
```

- [ ] **Step 2: Run and verify**

Run: `cd "/c/Users/jonas/Projects/AA Pilot Training" && python3 build_presentation.py`
Expected: `Saved: US Air Attack Pilot Orientation.pptx (18 slides)`

---

### Task 6: Section 5 — The Air Attack Role Reinforcement (Slides 19-23)

**Files:**
- Modify: `build_presentation.py` (insert before SAVE section)

- [ ] **Step 1: Add the Air Attack Role slides**

```python
# ============================================================
# SECTION 5: THE AIR ATTACK ROLE — REINFORCEMENT
# ============================================================

# Slide 19 — Your Job
add_content_slide(prs,
    title_text="Your Job",
    bullets=[
        {"text": "Fly the aircraft — provide a stable observation platform for the AAS"},
        {"text": "Maintain separation from terrain and traffic"},
        {"text": "Stay ahead of the mission — anticipate AAS needs, manage fuel and endurance"},
        {"text": "Systems, navigation, and comms management"},
    ],
    notes_text=(
        "Your primary responsibility is flying the airplane and providing a stable "
        "platform so the AAS can do their job. Maintain separation, stay ahead of the "
        "mission, and manage the systems. Anticipate what the AAS is going to need — "
        "don't wait to be asked. If you can see that fuel is going to be a factor, say "
        "something early. If weather is building, flag it. Be proactive."
    )
)

# Slide 20 — What You Don't Do
add_content_slide(prs,
    title_text="What You Don't Do",
    bullets=[
        {"text": "No transmissions on tactical frequencies — except safety of flight"},
        {"text": "The AAS handles all tactical radio, target descriptions, airtanker coordination"},
        {"text": "You monitor all frequencies and coordinate with AAS via intercom only"},
        {"text": "Exception: if AAS is task-saturated and a critical call comes in, you answer and brief AAS on intercom"},
    ],
    notes_text=(
        "You do not transmit on tactical frequencies unless it's a safety-of-flight "
        "issue. The AAS handles all of the tactical communication — target "
        "descriptions, airtanker coordination, everything. You listen, you monitor, "
        "and you coordinate with the AAS on intercom. The one exception: if the AAS "
        "is completely task-saturated — say they're heads down on FM with ground "
        "resources — and a critical call comes in on air-to-air, you can answer it "
        "and then brief the AAS on intercom."
    )
)

# Slide 21 — CRM in the Air Attack
add_content_slide(prs,
    title_text="CRM in the Air Attack",
    bullets=[
        {"text": "Monitor all frequencies, provide traffic callouts, assist during high workload"},
        {"text": "Communicate proactively: fuel state, weather changes, performance limitations, system anomalies"},
        {"text": "Withholding information erodes CRM — if something is developing, say it"},
        {"text": "Both crew members back each other up"},
    ],
    notes_text=(
        "Good CRM in the air attack means being a proactive partner. You're "
        "monitoring all the frequencies — if you hear traffic that the AAS might have "
        "missed, call it out on intercom. If fuel state is becoming a factor, don't "
        "wait until it's critical. If you see weather building or you've got a system "
        "anomaly, say something. Withholding information — even unintentionally, by "
        "just not speaking up — erodes the CRM relationship. You back each other up."
    )
)

# Slide 22 — Key Tactical Reminders
add_content_slide(prs,
    title_text="Key Tactical Reminders",
    bullets=[
        {"text": "1,000 ft vertical separation from dissimilar aircraft"},
        {"text": "Orbit at or above 1,500 ft AGL — hard floor"},
        {"text": "170 KIAS or less within FTA, 150 KIAS or less maneuvering"},
        {"text": "Overhead pattern entry preferred"},
        {"text": "FTA entry: communicate, clearance, comply — no entry without clearance"},
    ],
    notes_text=(
        "You covered all of this in detail during the ground school. These are the "
        "key numbers to keep in your head. 1,000 feet vertical separation. 1,500 AGL "
        "minimum orbit altitude. 170 knots or less in the FTA, 150 or less when "
        "maneuvering. Overhead pattern entry when practical. And never enter an FTA "
        "without clearance — communicate, get clearance, comply."
    )
)

# Slide 23 — Audio Panel Management
add_content_slide(prs,
    title_text="Audio Panel Management",
    bullets=[
        {"text": "You are monitoring multiple frequencies — use individual volume controls to manage them"},
        {"text": "Highest priority frequency (air-to-air) at highest volume"},
        {"text": "Lowest priority (dispatch) turned down — but never turned off"},
        {"text": "You are the backup for the ATGS — your situational awareness depends on hearing what's going on"},
        {"text": "Never turn a frequency off"},
    ],
    notes_text=(
        "When you're monitoring four frequencies at once, use the individual volume "
        "controls on the audio panel to manage them. Air-to-air is your highest "
        "priority — keep it loudest. Dispatch is probably your lowest priority — turn "
        "it down so it's not competing. But never turn a frequency off. You are the "
        "backup for the ATGS. Your collective role is managing the airspace and "
        "maintaining situational awareness. If you've got a radio turned off, you "
        "can't hear what's going on, and you can't back them up."
    )
)
```

- [ ] **Step 2: Run and verify**

Run: `cd "/c/Users/jonas/Projects/AA Pilot Training" && python3 build_presentation.py`
Expected: `Saved: US Air Attack Pilot Orientation.pptx (23 slides)`

---

### Task 7: Section 6 — Safety & Bright Lines (Slides 24-27)

**Files:**
- Modify: `build_presentation.py` (insert before SAVE section)

- [ ] **Step 1: Add the Safety and Bright Lines slides**

```python
# ============================================================
# SECTION 6: SAFETY & BRIGHT LINES
# ============================================================

# Slide 24 — Know Your Limits
add_content_slide(prs,
    title_text="Know Your Limits",
    bullets=[
        {"text": "Operate within your own envelope — not the airplane's, yours"},
        {"text": "If you're out of touch with your own ability, you don't know where the edge is"},
        {"text": "That's how people get killed in this business"},
        {"text": "Humility is not weakness — the pilot who thinks they're beyond learning is the one who gets in trouble"},
        {"text": "Aerial firefighting is high-risk — the mindset of honest self-assessment starts here, now"},
    ],
    notes_text=(
        "This is a direct message. The airplane has an operating envelope, and so do "
        "you. Yours is smaller. If a pilot is out of touch with their own ability — "
        "if they think they're hot shit and they're really not — that leads them to "
        "do things in an airplane that are beyond their skill level. And that's how "
        "people get killed in this business. Humility isn't weakness. It's what keeps "
        "you alive. The pilot who thinks they're beyond learning is the one who puts "
        "themselves in a situation they can't handle. That mindset starts here. Now. "
        "Be the pilot who knows what they don't know and keeps working to close the gap."
    )
)

# Slide 25 — Safety Culture
add_content_slide(prs,
    title_text="Safety Culture",
    bullets=[
        {"text": "Safety is never compromised for the mission"},
        {"text": "Coulson backs you when you make the safe call — no penalty, no questions"},
        {"text": "Proactive risk mitigation, not reactive"},
        {"text": "Report through SMS — that's how we learn and improve"},
    ],
    notes_text=(
        "If you make a decision based on safety, we back you. No penalty, no "
        "questions. We would rather you make a conservative call and have a "
        "conversation about it than push into something that doesn't feel right. "
        "Report through SMS when something happens — near misses, safety concerns, "
        "anything. That's how we learn and how we get better as a company."
    )
)

# Slide 26 — Bright Lines
add_content_slide(prs,
    title_text="Bright Lines",
    bullets=[
        {"text": "We've hired good people and we trust you — but there are lines that can't be crossed:"},
        {"text": "Integrity — honesty matters. If something happened, tell us. Lying about it is worse than the thing itself.", "level": 1},
        {"text": "Willful disregard for standard operating procedures", "level": 1},
        {"text": "Representing Coulson in a way that damages the company's reputation", "level": 1},
        {"text": "Conduct that undermines trust — with your crew, with agency personnel, with the team", "level": 1},
        {"text": "This isn't about mistakes — everyone makes mistakes and we'll work through them. This is about character."},
    ],
    notes_text=(
        "We've done well hiring good people and I'm not standing up here lecturing "
        "you about basic decency. But I do want to be clear about where the lines are. "
        "Integrity is non-negotiable. If something happened — you made a mistake, "
        "something went wrong — tell us. We'll work through it. Lying about it is "
        "worse than whatever the thing was. Willful disregard for SOPs, representing "
        "the company poorly, conduct that undermines the trust of the people you work "
        "with — those are bright lines. This isn't about operational mistakes. Everyone "
        "makes those and we'll work through them together. This is about character."
    )
)

# Slide 27 — FRAT and Tactical Pause
add_content_slide(prs,
    title_text="FRAT and Tactical Pause",
    bullets=[
        {"text": "FRAT scoring: Green (proceed), Yellow (mitigation required), Red (not authorized)"},
        {"text": "Tactical pause: either crew member can call it, at any time, no penalty"},
        {"text": "If something doesn't feel right, say so"},
        {"text": "We'd rather have a conversation than an incident"},
    ],
    notes_text=(
        "The FRAT gives you a structured way to assess risk. Green means go. Yellow "
        "means you need to identify and apply mitigations before proceeding. Red means "
        "the mission is not authorized. And at any point during a mission, either crew "
        "member can call a tactical pause — no penalty, no judgment. If something "
        "doesn't feel right, say so. We would rather have a conversation about it "
        "than deal with an incident."
    )
)
```

- [ ] **Step 2: Run and verify**

Run: `cd "/c/Users/jonas/Projects/AA Pilot Training" && python3 build_presentation.py`
Expected: `Saved: US Air Attack Pilot Orientation.pptx (27 slides)`

---

### Task 8: Sections 7 & 8 — What You're Checked On + Close (Slides 28-30)

**Files:**
- Modify: `build_presentation.py` (insert before SAVE section)

- [ ] **Step 1: Add the final slides**

```python
# ============================================================
# SECTION 7: WHAT YOU'RE BEING CHECKED ON
# ============================================================

# Slide 28 — Training Form (TRN-04)
add_content_slide(prs,
    title_text="Training Form (TRN-04)",
    bullets=[
        {"text": "Ground school + flight training competency areas"},
        {"text": "What you'll be signed off on before your OPC:"},
        {"text": "Regulatory knowledge and company procedures", "level": 1},
        {"text": "Mission planning and dispatch", "level": 1},
        {"text": "Fire operations knowledge", "level": 1},
        {"text": "Aircraft handling", "level": 1},
        {"text": "Single-pilot operations and resource management", "level": 1},
    ],
    notes_text=(
        "The training form covers everything you need to be signed off on before "
        "your OPC. It's organized into ground school competencies and flight training "
        "competencies. Your training pilot will work through this with you and sign "
        "off each area as you demonstrate proficiency."
    )
)

# Slide 29 — OPC Checkride (OPC-04)
add_content_slide(prs,
    title_text="OPC Checkride (OPC-04)",
    bullets=[
        {"text": "Approximately 1-hour evaluation flight — check pilot acts as AAS"},
        {"text": "26 evaluated items across 6 competency areas"},
        {"text": "Automatic failure gates: safety event, repeated deficiency, checklist discipline"},
        {"text": "Outcomes: Approved or Not Approved"},
        {"text": "Annual recurrency required"},
    ],
    notes_text=(
        "The OPC is about a one-hour evaluation flight where the check pilot flies "
        "in the right seat acting as the AAS. You'll be evaluated on 26 items across "
        "six competency areas — everything from mission knowledge and FTA procedures "
        "to non-technical skills like CRM and decision-making. There are automatic "
        "failure gates: a safety event, a repeated deficiency, or checklist discipline "
        "failures will result in a Not Approved outcome. The OPC is required annually "
        "to maintain your qualification."
    )
)

# ============================================================
# SECTION 8: CLOSE
# ============================================================

# Slide 30 — Chief Pilot's Message
add_title_slide(prs,
    title_text="Chief Pilot's Message",
    subtitle_text="Questions",
    notes_text=(
        "You're part of something special at Coulson. We're investing in you and we "
        "take that seriously. The air attack seat is the best seat in the house to "
        "learn aerial firefighting — use it for what it's worth. Perform well, learn "
        "constantly, represent Coulson with pride. The tanker seat is earned, not "
        "given. This is where you earn it.\n\n"
        "Open the floor for questions."
    )
)
```

- [ ] **Step 2: Run and verify final build**

Run: `cd "/c/Users/jonas/Projects/AA Pilot Training" && python3 build_presentation.py`
Expected: `Saved: US Air Attack Pilot Orientation.pptx (30 slides)`

---

### Task 9: Final verification and cleanup

**Files:**
- Read: `US Air Attack Pilot Orientation.pptx`

- [ ] **Step 1: Run a verification script to confirm all 30 slides have content and notes**

Run:
```bash
cd "/c/Users/jonas/Projects/AA Pilot Training" && python3 -c "
from pptx import Presentation
prs = Presentation('US Air Attack Pilot Orientation.pptx')
print(f'Total slides: {len(prs.slides)}')
print()
for i, slide in enumerate(prs.slides, 1):
    title = slide.shapes.title.text if slide.shapes.title else '(no title)'
    has_notes = bool(slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip())
    layout = slide.slide_layout.name
    print(f'Slide {i:2d} [{layout:20s}] {title:40s} notes={has_notes}')
"
```

Expected: 30 slides listed, all with `notes=True`, correct titles matching the spec.

- [ ] **Step 2: Open the PPTX and visually verify**

Open `US Air Attack Pilot Orientation.pptx` in PowerPoint and check:
- Coulson Carbon theme applied (dark background, angled bands)
- Garamond font throughout
- Logo on title slide
- Bullet indentation working (level 0 and level 1)
- Speaker notes visible in Notes view
- All 30 slides present with correct content

- [ ] **Step 3: Clean up if needed**

If any visual issues are found (font size, spacing, alignment), adjust the `build_presentation.py` script and re-run.
