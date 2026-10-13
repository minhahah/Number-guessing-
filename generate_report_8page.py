from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
import datetime, os

BASE = r"E:\minha_java"
OUT = os.path.join(BASE, "NumberGuessingGame_Report_Academic.docx")
NAVY = RGBColor(0x1F, 0x3A, 0x5F)
ACCENT = RGBColor(0x2D, 0x5A, 0xA0)

doc = Document()

# --- page setup: narrow margins to hold 8 pages ---
for s in doc.sections:
    s.top_margin = Inches(0.5)
    s.bottom_margin = Inches(0.5)
    s.left_margin = Inches(0.6)
    s.right_margin = Inches(0.6)

style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(10.5)
style.paragraph_format.space_after = Pt(4)
style.paragraph_format.space_before = Pt(0)
style.paragraph_format.line_spacing = 1.0

for lvl in ('Heading 1', 'Heading 2', 'Heading 3'):
    st = doc.styles[lvl]
    st.paragraph_format.space_before = Pt(8)
    st.paragraph_format.space_after = Pt(3)

def h1(t):
    h = doc.add_heading(t, level=1)
    for r in h.runs:
        r.font.color.rgb = NAVY; r.font.size = Pt(14)
    return h

def h2(t):
    h = doc.add_heading(t, level=2)
    for r in h.runs:
        r.font.color.rgb = ACCENT; r.font.size = Pt(12)
    return h

def para(text, italic=False, size=None, bold=False, align=None):
    p = doc.add_paragraph()
    if align: p.alignment = align
    r = p.add_run(text)
    r.bold = bold; r.italic = italic
    if size: r.font.size = Pt(size)
    return p

def bullets(items):
    for it in items:
        if isinstance(it, tuple):
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(it[0]); r.bold = True
            p.add_run(it[1])
        else:
            p = doc.add_paragraph(it, style='List Bullet')
            p.paragraph_format.space_after = Pt(2)

def table(headers, rows, fs=9):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Light Grid Accent 1'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, hh in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = ''
        r = c.paragraphs[0].add_run(hh)
        r.bold = True; r.font.size = Pt(fs)
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = ''
            r = cells[i].paragraphs[0].add_run(str(val))
            r.font.size = Pt(fs)
    sp = doc.add_paragraph('')
    sp.paragraph_format.space_after = Pt(2)
    sp.paragraph_format.space_before = Pt(0)
    return t

def code(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text)
    r.font.name = 'Consolas'; r.font.size = Pt(8)
    pPr = p._p.get_or_add_pPr()
    shd = pPr.makeelement(qn('w:shd'), {qn('w:fill'): 'F2F2F2', qn('w:val'): 'clear'})
    pPr.append(shd)
    return p

def shot(path, caption):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(1); p.paragraph_format.space_before = Pt(1)
    p.add_run().add_picture(path, width=Inches(4.9))
    c = doc.add_paragraph(); c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c.paragraph_format.space_after = Pt(3)
    r = c.add_run(caption)
    r.italic = True; r.font.size = Pt(8.5); r.font.color.rgb = RGBColor(0x55,0x55,0x55)

# ================= COVER + 1. INDEX (Page 1) =================
t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run('Java Swing GUI-Based\nNumber Guessing Game'); r.bold = True; r.font.size = Pt(24); r.font.color.rgb = NAVY
s = doc.add_paragraph(); s.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = s.add_run('Academic Project Report  •  JProgressBar Edition'); r.font.size = Pt(12); r.font.color.rgb = ACCENT
i = doc.add_paragraph(); i.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = i.add_run(f"{datetime.date.today().isoformat()}  •  NumberGuessingGame.java  •  Range 1–100  •  10 Attempts"); r.font.size = Pt(9)

h1('1. Index Page')
para('Section listing with page-number placeholders (finalise after PDF pagination).', italic=True, size=9)
table(['No.', 'Section', 'Content', 'Page'],
[
 ['1', 'Index Page', 'Section listing + placeholders', '1'],
 ['2', 'Project Overview', '2.1 Introduction / 2.2 Objectives / 2.3 Audience', '2'],
 ['3', 'Technical Specifications', '3.1 Swing / 3.2 AWT + Events / 3.3 Random', '3'],
 ['4', 'Core Logic & Flow', '4.1 Init / 4.2 Input pipeline / 4.3 Comparison engine', '4'],
 ['5', 'UI Interaction & Screenshots', 'Placeholders 1–5 + layout tables + logic', '5–7'],
 ['6', 'Conclusion & Future Scope', '6.1 Stability / 6.2 Extensions', '8'],
])

# ================= 2. OVERVIEW (Page 2) =================
h1('2. Project Overview')
h2('2.1 Introduction')
para('The Number Guessing Game is an event-driven desktop micro-application built with Java Swing. The system conceals an integer in [1, 100] and grants ten attempts to identify it. Each valid submission returns a directional hint (“Too High” / “Too Low”), depletes one tracking unit, and synchronises a JLabel counter with a string-painted JProgressBar whose fill and colour encode residual state. A single JFrame is partitioned into header, central stack (input, feedback, counter, bar) and Reset controls. All interaction flows through an ActionListener on the ‘Guess’ button, with lifecycle control via ‘Reset Game’. The project is a compact, dependency-free demonstration of state management, rule processing, layout composition, and defensive parsing.')
para('Gameplay is deliberately simple so that attention remains on engineering quality: every guess exercises the same deterministic pipeline — read the text stream, enforce validation boundaries, decrement the reserve, evaluate the comparison rules, and repaint all dependent components atomically on the Event Dispatch Thread. The quantitative reserve is therefore never shown in two conflicting ways; the textual label and the visual bar are different renderings of the single attemptsLeft variable. Colour carries urgency semantics throughout: blue for information, orange for guidance, red for errors and danger, and green for success and health. This consistent visual grammar lets a first-time player infer game state without reading documentation.')
h2('2.2 Project Objectives')
bullets([
 ('Event-driven UI: ', 'capture button clicks via ActionListener.actionPerformed(ActionEvent) and route them to deterministic logic.'),
 ('State management: ', 'couple targetNumber + attemptsLeft to label, progress bar, and button enablement across the lifecycle.'),
 ('Operational rule processing: ', 'encode match / high / low, per-guess depletion, and win-on-equality / loss-on-zero terminal rules.'),
 ('Layout configuration: ', 'compose BorderLayout + GridLayout (4×1) + FlowLayout with sizing, borders, fonts, colours.'),
 ('Error boundaries: ', 'quarantine empty, non-numeric (NumberFormatException), and out-of-range input before any state mutation.'),
 ('Visual state: ', 'replace a text-only counter with a proportional, colour-graded JProgressBar (green / orange / red).'),
])
h2('2.3 Target Audience')
para('Designed as an engaging educational micro-app for introductory Java/OOP learners (single-class readability), HCI students (layouts, feedback semantics, focus handling), and instructors needing a deterministic, easily testable lab/viva artefact. Only a standard JDK is required.')
para('For learners, the code rewards close reading: the constructor shows component assembly order, startNewGame() shows state initialisation, handleGuess() shows guarded rule evaluation, and updateAttemptsBar() shows derived-view synchronisation. For instructors, every behaviour is reproducible — the random bound, the validation rejections, the colour thresholds, and the terminal freezes can each be demonstrated live in under a minute. For self-learners, the build-run-modify loop is frictionless: one file, one compile command, zero dependencies, and immediate visual feedback for every edit.')

# ================= 3. TECH SPECS (Page 3) =================
h1('3. Technical Specifications & Libraries Used')
table(['Category', 'Specification'],
[
 ['Language / Toolkit', 'Java SE 8+ (tested 21.0.8 LTS); Swing over AWT; no external deps'],
 ['Window', 'JFrame 400×300, BorderLayout, EXIT_ON_CLOSE, centred; GridLayout(4,1) centre'],
 ['Build / Run', 'javac NumberGuessingGame.java  •  java NumberGuessingGame (on EDT)'],
])
h2('3.1 javax.swing.* — Containers and Active Components')
para('Import javax.swing.* builds every visible element:', italic=True, size=9)
para('Swing provides lightweight components rendered in Java rather than by the native windowing system, which guarantees identical appearance across platforms. The game exploits this by relying on Swing borders, fonts, and colours for its entire visual identity. Because the class extends JFrame, the frame itself is the application object — construction, layout, listener registration, and initial-state setup all occur in one place, keeping the component graph easy to trace.')
table(['Component', 'Why Imported'],
[
 ['JFrame', 'Top-level window; extended by the game class; hosts title, sizing, layout, visibility.'],
 ['JPanel', 'Grouping containers: centre stack, input row, bottom Reset row + EmptyBorder padding.'],
 ['JLabel', 'Display: title, “Enter your guess” prompt, colour-coded feedbackLabel, attemptsLabel counter.'],
 ['JTextField (6 cols)', 'Input stream via getText(); cleared + refocused per round; locked at terminal states.'],
 ['JButton', '“Guess” fires handleGuess(); “Reset Game” fires startNewGame(); enablement enforces lifecycle.'],
 ['JProgressBar (0–10)', 'Visual tracker: max 10, init full, string-painted “N / 10 Attempts Left”, green/orange/red fill.'],
 ['JOptionPane*', 'Cited family member; terminal text kept inline here, dialog is a drop-in alternative.'],
])
h2('3.2 java.awt.* & java.awt.event.* — Layout, Style, Events')
para('The AWT layer supplies geometry and the delegation event model on which Swing is built. Layout managers remove the need for absolute pixel positioning: the frame adapts to font metrics and border insets automatically, while FlowLayout rows wrap gracefully if the window is resized. The event package implements the observer pattern — controls do not poll for input; instead, the runtime dispatches ActionEvent objects to registered ActionListener implementations.', italic=True, size=9)
table(['Element', 'Why Imported'],
[
 ['BorderLayout / GridLayout / FlowLayout', 'Frame regions; 4-row centre stack (input, feedback, counter, bar); button rows.'],
 ['Font / Color', 'Arial hierarchy (16/14/13/12pt); BLUE info, ORANGE hint, RED error/danger, Green success.'],
 ['ActionListener / ActionEvent', 'Event contract + event object; getSource() dispatches Guess vs Reset.'],
])
h2('3.3 java.util.Random — Pseudorandom Engine 1–100')
code('Random rand = new Random();\ntargetNumber = rand.nextInt(100) + 1; // uniform [1,100]')
para('nextInt(100) yields 0–99; +1 shifts to the advertised 1–100 range. Invoked once per game inside startNewGame() (constructor + every Reset), so each round conceals a fresh value with no extra entropy source.')
para('Using Random rather than Math.random() is a deliberate readability choice: the generator object is explicit, the bound arithmetic is visible in a single expression, and seeding can later be introduced for deterministic testing (e.g. new Random(42)) without changing any other line. The engine easily satisfies game unpredictability — with 100 equiprobable outcomes, brute-force guessing within ten attempts succeeds only by luck, which is precisely the intended challenge.')

# ================= 4. CORE LOGIC (Page 4) =================
h1('4. Core Application Logic & Component Flow')
h2('4.1 Game Initialization')
para('Initialisation order matters: components must exist before startNewGame() touches them, listeners must be registered before the window is shown, and the window must be sized and centred before becoming visible. The constructor enforces this sequence — build header, build centre rows including the progress bar, build Reset controls, register this as listener for both buttons, initialise state, then setVisible(true). Any deviation (e.g. showing the frame before resetting the bar) would flash a transient incorrect state.', italic=True, size=9)
bullets([
 ('Launch: ', 'main() → SwingUtilities.invokeLater(() -> new NumberGuessingGame()) on the EDT; constructor builds UI, wires listeners, calls startNewGame().'),
 ('startNewGame(): ', 'targetNumber = nextInt(100)+1; attemptsLeft = 10; field cleared/editable; Guess enabled, Reset disabled; BLUE “new number” message; counter “10”; bar reset to 10/10 green. Centred via setLocationRelativeTo(null).'),
])
h2('4.2 Input Processing Pipeline')
para('handleGuess() validates before mutating — failures return early with a red message and zero attempt cost:', italic=True, size=9)
para('The pipeline ordering is a security property, not merely a convenience. The empty check runs first because trim() on a blank field would otherwise reach the parser and throw; the parse runs second because range comparison on a non-integer is meaningless; the range check runs last because only a genuine integer can be judged against [1, 100]. Each guard is an independent if/try block with its own message, so the player always learns exactly which boundary was violated and can correct precisely that aspect.')
table(['Stage', 'Check', 'Failure Message (Red, No Cost)'],
[
 ['1. Read', 'getText().trim()', '—'],
 ['2. Empty', 'isEmpty()', '“Please enter a valid number!”'],
 ['3. Parse', 'parseInt in try/catch NumberFormatException', '“Invalid characters! Enter digits only.”'],
 ['4. Range', '<1 or >100', '“Out of range! Guess must be 1 to 100.”'],
 ['5. Commit', 'attemptsLeft-- + sync label/bar', 'Proceed to comparison'],
])
h2('4.3 Comparison Engine')
table(['Rule', 'Condition → Response'],
[
 ['WIN', 'guess == target → green “Correct! You won!” + endGame() freeze.'],
 ['LOSS', 'attemptsLeft == 0 → red “Game Over! Number was: N”, 0/10 red bar + endGame().'],
 ['HINT Low/High', 'guess < / > target → orange “Too Low / Too High”, field cleared + refocus.'],
])
code('attemptsLeft--; attemptsLabel.setText(...); updateAttemptsBar(); // green >5, orange 3–5, red ≤2\n'
     'if (guess==target) win; else if (attemptsLeft==0) loss+reveal; else hint;')
para('Branch order is load-bearing: the win test precedes the depletion test so that a correct final-attempt guess counts as a victory rather than a loss, and both precede the hint branch so terminal states never also emit guidance. Post-round housekeeping — guessField.setText("") plus requestFocus() — keeps the interaction loop tight: the player never needs the mouse to continue. The bar colour law (green above half, orange at half, red at two or fewer) converts the raw integer into an at-a-glance urgency signal that reinforces the numeric label.')

# ================= 5. UI + SCREENSHOTS (Pages 5–7) =================
h1('5. User Interface Interaction & Screenshot Breakdown')
para('Each sub-section below gives a layout table, a screenshot figure with caption, and click-action logic.', italic=True, size=9)

h2('5.1 Placeholder 1 — Initial State Launch Screen')
table(['Region', 'State'],
[
 ['NORTH', 'Title “Guess a number between 1 and 100” (Bold 16).'],
 ['Centre R1–R4', 'Empty field + armed Guess; BLUE “new number” msg; “10”; bar 10/10 green “10 / 10 Attempts Left”.'],
 ['SOUTH / Frame', 'Reset DISABLED; 400×300 centred via setLocationRelativeTo(null) on EDT.'],
], fs=8.5)
shot(os.path.join(BASE,'screenshot_1_start.png'), 'Ref. Fig. 5.1 — Initial: 10/10 green, Reset disabled.')
para('Logic: this is the post-condition of startNewGame(); the concealed target has just been drawn, attemptsLeft = 10, and no click has occurred. The empty editable field plus armed Guess button invite the first entry, while the disabled Reset button communicates that there is nothing yet to restart. Text and visual counters agree at maximum, establishing the baseline against which every later depletion is judged. Centring via setLocationRelativeTo(null) guarantees this first impression is identical on any display.', size=9)

h2('5.2 Placeholder 2 — Active Gameplay Input')
table(['Region', 'State'],
[
 ['Input', 'Number typed (e.g. “78”); Guess click → ActionEvent → handleGuess().'],
 ['Feedback/Counter/Bar', 'ORANGE “Too High! Try lower.”; counter −1; bar drops smoothly with new “N / 10” text + recoloured fill.'],
 ['Controls', 'Reset still disabled; focus returns via requestFocus(); exactly one unit per valid click.'],
], fs=8.5)
shot(os.path.join(BASE,'screenshot_2_progress_green.png'), 'Ref. Fig. 5.2a — Healthy play: “25”, 7/10 green, Too Low hint.')
shot(os.path.join(BASE,'screenshot_3_warning_orange.png'), 'Ref. Fig. 5.2b — Warning play: “78”, 4/10 orange, Too High hint.')
para('Logic: read → validate (pass) → decrement → synchronise label and bar → hint branch. The setValue() call shortens bar fill proportionally, setString() rewrites the embedded “N / 10” text, and setForeground() re-evaluates the colour threshold — all inside one EDT dispatch, so the change appears instantaneous. Orange is reserved for non-terminal guidance, cleanly separated from red errors and green success, and the cleared, refocused field keeps play keyboard-driven.', size=9)

h2('5.3 Placeholder 3 — Invalid Input / Error Boundary')
table(['Region', 'State'],
[
 ['Input', 'Symbols retained, e.g. “ab@#”; Guess clicked.'],
 ['Feedback', 'RED “Invalid characters! Enter digits only.” (NumberFormatException path).'],
 ['Counter/Bar', 'FROZEN (still 10/10 green) — early return precedes depletion; loop unbroken.'],
], fs=8.5)
shot(os.path.join(BASE,'screenshot_7_invalid.png'), 'Ref. Fig. 5.3 — Error boundary: “ab@#”, red error, bar frozen 10/10.')
para('Logic: Integer.parseInt() throws NumberFormatException; the catch block paints the feedback red and returns before attemptsLeft-- is reachable. Empty-string and out-of-range guards behave identically. Because mutation is sequenced strictly after validation, hostile or accidental input can never drive the reserve negative or desynchronise label from bar — the evaluator may click Guess repeatedly with invalid data and observe zero bar movement, confirming loop integrity.', size=9)

h2('5.4 Placeholder 4 — Win State Terminal Screen')
table(['Region', 'State'],
[
 ['Feedback', 'GREEN “Correct! You won the game!” (JOptionPane dialog substitutable).'],
 ['Inputs', 'Field non-editable; Guess disabled — canvas frozen; bar frozen at residual value.'],
 ['Controls', 'Reset ENABLED — sole action, re-invokes startNewGame().'],
], fs=8.5)
shot(os.path.join(BASE,'screenshot_6_win.png'), 'Ref. Fig. 5.4 — Win: green message, locked inputs, Reset enabled.')
para('Logic: the winning click passes validation, consumes its unit, and satisfies guess == target on the first branch — loss and hint branches are skipped. endGame() then triple-toggles the interface (field non-editable, Guess disabled, Reset enabled), producing an unambiguous stable terminal screen. Because javax.swing.* is already imported, a modal JOptionPane.showMessageDialog() congratulation can be added beside the inline label with zero new dependencies.', size=9)

h2('5.5 Placeholder 5 — Loss State Terminal Screen')
table(['Region', 'State'],
[
 ['Feedback', 'RED “Game Over! The number was: N” — solution disclosure.'],
 ['Counter/Bar', '“0 remaining”; bar 0/10 empty RED via updateAttemptsBar().'],
 ['Controls', 'Same freeze as win; Reset ENABLED — full restart without JVM relaunch.'],
], fs=8.5)
shot(os.path.join(BASE,'screenshot_5_gameover.png'), 'Ref. Fig. 5.5 — Loss: 0/10 red, solution revealed, Reset enabled.')
para('Logic: the final valid click decrements the reserve to exactly zero, fails the win test, and satisfies the depletion rule — revealing the target concealed all round. Exhaustion is communicated on four channels at once: quantitatively (0/10), textually (“0 remaining”), chromatically (red), and procedurally (freeze plus a single recovery action). Clicking Reset re-invokes startNewGame(): fresh bound, full reserve, restored controls — a complete lifecycle restart without relaunching the JVM.', size=9)

# ================= 6. CONCLUSION (Page 8) =================
h1('6. Conclusion & Future Scope')
h2('6.1 Structural Stability Summary')
para('The UI framework is highly stable: a three-region single frame isolates presentation (labels/bar), control (two buttons), and state (two integers). All mutations funnel through startNewGame() (birth), handleGuess() (evolution), and endGame() (termination); label and bar update atomically, enablement toggles only at boundaries, and invalid input is quarantined pre-mutation. The bar adds proportional, colour-graded feedback with no new threads, listeners, or dependencies — compiling on any standard JDK and degrading gracefully under all input classes.')
para('Stability is verifiable by inspection and by test: rapid Guess clicks cannot skip the validation gate; invalid strings cannot move the bar; the win-on-last-attempt ordering is encoded in branch sequence; and Reset restores every component including the bar to its exact launch condition. The report screenshots themselves constitute evidence — each figure shows label, bar value, bar colour, message colour, and button enablement in mutual agreement, with no reachable state in which they contradict.')
h2('6.2 Future Evolutionary Paths')
para('The extensions below are ordered by complexity; each builds on the current architecture without disturbing the validated event loop or error boundaries:', italic=True, size=9)
bullets([
 ('Persistent high scores: ', 'append wins (date, attempts, target) to a text/CSV log via java.io/nio; show best score and a “View Scores” reader.'),
 ('Audio feedback: ', 'short cues via javax.sound.sampled (win chime, error buzz) with mute toggle — multimodal feedback.'),
 ('Difficulty toggles: ', 'Easy 1–50/15, Normal 1–100/10, Hard 1–200/6 — re-parameterise bound, MAX_ATTEMPTS, bar max, and colour thresholds at reset.'),
])
para('Build: cd /d E:\\minha_java  •  javac NumberGuessingGame.java  •  java NumberGuessingGame. Screenshots 600×420 rendered from identical bar logic.', size=8.5, italic=True)
para('Verification matrix (all confirmed in this workspace):', bold=True, size=9.5)
table(['Test', 'Input / Action', 'Expected Result'],
[
 ['Compile', 'javac NumberGuessingGame.java', 'Clean build, no errors'],
 ['Launch', 'java NumberGuessingGame', 'Centred 400×300 window, 10/10 green bar'],
 ['Invalid', '“ab@#” + Guess', 'Red format error; bar frozen'],
 ['Hint', 'Wrong number + Guess', 'Orange hint; reserve −1; bar drops + recolour'],
 ['Win', 'Guess == target', 'Green win; inputs locked; Reset on'],
 ['Loss + Reset', '10 misses, then Reset', 'Red 0/10 + solution; Reset restores 10/10 green'],
], fs=8.5)

doc.save(OUT)
print('Saved', OUT)
