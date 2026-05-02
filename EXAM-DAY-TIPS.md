# MO-211 Exam Day Tips 🎯

Knowing the material is half the battle — knowing the **test environment** is the other half. The MO-211 exam is a live, performance-based exam, not multiple choice. You drive a real (sandboxed) copy of Excel and earn points by completing tasks correctly. This guide walks through what to expect and how to spend your 50 minutes wisely.

---

## 🖥 What the exam looks like

* A **live, sandboxed instance of Excel** runs inside the test delivery window. It behaves like the real app — same ribbon, same dialogs, same shortcuts.
* The exam contains **one or more "Projects"**, each a downloadable workbook with multiple tasks. In total you face **~35 tasks** across all projects.
* A **task pane** at the bottom (or side) of the screen shows the current task description, a task list, and navigation buttons.
* The **50-minute clock** is always visible. It does not pause.
* Three buttons matter most:
  * **Mark for Review** — flag a task to come back to later.
  * **Reset Project** — discard all changes to the current project workbook and start it over. Use only as a last resort.
  * **Restart Section / Submit** — finalize the project. After submitting a project you cannot return to it.

---

## 📅 Before exam day

* **Confirm your delivery method.** Either an in-person Certiport testing center or an **online proctored** session through Certiport's Compass platform. Online proctoring requires a webcam, a clean desk, and a single monitor — verify your setup the day before, not five minutes before.
* **Practice on something that matches the real UI.** [GMetrix](https://www.gmetrix.com/) and **CertPREP** both ship MO-211 practice tests that closely mimic the live test interface. Doing one full practice run is worth ten YouTube videos.
* **Run the course's `*.test-yourself.md` files at least twice.** First pass: with notes. Second pass: closed-book and timed. Any task you stumble on twice is a topic to review the next morning.
* **Time yourself on a mock exam.** Aim to finish with **5+ minutes left** so you have a buffer to revisit your *Mark for Review* items.
* **Sleep.** A tired brain misreads task prompts, and misreading a prompt is the single fastest way to lose points.

---

## ⏱ During the exam

* **Read the entire task before clicking anything.** Tasks often contain two or three sub-requirements in one paragraph (e.g. "create a PivotTable on a new sheet named *Summary*, group dates by quarter, and show values as % of grand total"). Missing a sub-requirement loses you the whole task's points.
* **Use Mark for Review aggressively.** If a task takes more than ~90 seconds and you are not close to done, mark it and move on. Banking the easy points first beats dying on one hard task.
* **Save often** with `Ctrl+S`. The test environment is mostly stable, but a crash on an unsaved workbook is recoverable — a crash with no saves is not.
* **Use Reset Project sparingly.** It is genuinely destructive in some delivery environments and will undo every change to that project's workbook, including completed tasks. Prefer `Ctrl+Z` first.
* **Answer literally what the task asks.** Do not "improve" the workbook. If the task says *change the chart title to "Q3 Revenue"*, change only the chart title. Anything else risks side-effect penalties (see below).
* **Trust your shortcuts.** See [`KEYBOARD-SHORTCUTS.md`](./KEYBOARD-SHORTCUTS.md). Every ribbon click is a few seconds you cannot get back.

---

## ⚠️ Common pitfalls

| Pitfall | What goes wrong | How to avoid it |
| --- | --- | --- |
| **Wrong file extension on save** | Saving a `.xlsm` (macro-enabled) workbook as `.xlsx` silently strips all macros — and the auto-grader sees zero points for the macro task | Always check the *Save as type* dropdown. If you wrote/recorded a macro, the file must stay `.xlsm`. |
| **Modifying cells outside the task's scope** | The auto-grader compares your workbook to a reference; unrelated edits count as side effects and dock points | After each task, undo any stray edits before moving on. `Ctrl+Z` is your friend. |
| **Approximate-match `VLOOKUP`** | Using `TRUE` (or omitting the 4th arg) when an exact match is required gives wrong answers | Default to `FALSE` for the 4th argument unless the task explicitly says "find the closest match" or "lookup in a sorted band/grade table" |
| **Waterfall chart missing "Set as Total"** | The final/total bar floats instead of resting on the axis, and the chart is marked wrong | Right-click the total data point → *Set as Total*. Same for any subtotal columns. |
| **PivotTable: changing field placement when only formatting was asked** | Task says "format the Sales values as Currency"; you also drag a field around → side-effect penalty | Re-read the task. If it says **format**, only format. Do not rearrange. |
| **Forgetting to confirm dialogs** | Editing a Name in Name Manager, or a data validation rule, without clicking *OK*/*Close* — changes are not committed | Always close dialogs explicitly before moving on. |
| **Hard-coding values that should be formulas** | Typing `42` instead of `=SUM(...)` works for the displayed value but the grader checks the formula | If the task says "use a function", use the function. |

---

## 🎒 What to bring / set up

**In-person testing center:**

* Valid government-issued **photo ID** (the name must match your Certiport profile exactly).
* Arrive 15 minutes early to check in.
* No phone, no smartwatch, no notes, no water bottle past the desk (lockers are usually provided).

**Online proctored:**

* The same photo ID, held up to the webcam during check-in.
* A **single, clear desk** — the proctor will ask you to pan the webcam around the room.
* **Disable additional monitors.** Online proctoring requires exactly one display.
* No phone within reach. No second person in the room.
* A glass of water is usually allowed; confirm with the proctor at check-in.
* A stable wired internet connection if you have one.

---

## 🏁 After the exam

* Your **score appears immediately** on screen after you submit. The passing threshold is **700 out of 1000**.
* You will receive a **digital badge via Credly** within 24 hours (check the email you registered with Certiport).
* Your **PDF certificate** is downloadable from your Certiport dashboard once the badge is issued.
* If you did not pass, Certiport's retake policy applies — typically a 24-hour wait before you can retake, and a small number of attempts on a single exam voucher.
* Update your LinkedIn — the certification is genuinely recognized by employers, and it is the whole reason you spent the time.

Good luck. You have done the work. Trust it.
