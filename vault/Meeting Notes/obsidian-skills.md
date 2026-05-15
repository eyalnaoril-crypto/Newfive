# Obsidian Skills — 3 סקילים לעבודה עם Obsidian Vault

## Overview

שלושת סקילי Obsidian שהותקנו בפרויקט. הקריטי שבהם — **`obsidian-vault-workflow`** — הוא המקור-אמת לאיך לעבוד עם ה-vault הזה (קריאה לפני משימה, כתיבה אחרי, מבנה Topic File, _index.md).

**שייכות:** כלל-פרויקטיים. **הסקיל `obsidian-vault-workflow` הוא חובה לכל סשן** (לפי החלטת גלעד — להפעיל אותו אוטומטית בתחילת כל פקודה).

**מיקום:** `.claude/skills/<skill-name>/SKILL.md`.

## Open Questions

- האם נצטרך גם `obsidian-bases` בפועל? כרגע אין `.base` files בפרויקט.
- האם להגדיר תקן מילוי שדה `## Open Questions` (להוסיף `- none` כשאין שאלות פתוחות, או להשמיט את הסעיף)?

## רשימת הסקילים

| סקיל | תפקיד |
|------|--------|
| **obsidian-vault-workflow** | **חובה — בתחילת ובסוף כל משימה.** קובע מבנה Topic File (Overview + Open Questions + Session Log), חוקי `_index.md`, סטטוס תגיות (`[shipped]`/`[wip]`/`[planned]`...), wikilinks. |
| **obsidian-markdown** | בעבודה על `.md` עם תחביר Obsidian — wikilinks, embeds, callouts, frontmatter, tags. |
| **obsidian-bases** | בעבודה על `.base` files (תצוגות database-like). כרגע לא בשימוש. |

## כללי המבנה שאוכפים על ה-vault הזה

מתוך `obsidian-vault-workflow`:

1. **קובץ אחד פר נושא** — לא קבצים מתוארכים נפרדים.
2. **Overview + Open Questions + Session Log** — שלושה סעיפים חובה בכל קובץ.
3. **Session Log = יומן כרונולוגי** — entries חדשים מתווספים בתחתית, כל אחד עם `### YYYY-MM-DD — title [status]`.
4. **`_index.md` בכל תיקייה** — מציג את כל קבצי הנושא בתיקייה.
5. **Wikilinks חובה** — `[[file-name]]` (בלי `.md`), לא markdown links.
6. **Status tags:** `[shipped]`, `[wip]`, `[planned]`, `[spiked]`, `[reverted]`, `[debug]`.

## קבצים קשורים

- **מנהל את כל ה-vault הזה:** `.claude/skills/obsidian-vault-workflow/SKILL.md`.
- **סקיל אחיותי:** [[superpowers-plugin]] (סקילי תהליך פיתוח).
- **כל קובצי ה-vault הקיימים** — מצייתים לכללים האלה: [[project-architecture]], [[root-claude-md]], [[agent-gal]], [[agent-irit]], [[agent-bar]], [[agent-barak]], [[root-readme]], [[root-env-config]], [[root-gitignore]], [[root-package-json]], [[claude-settings-local]].

## Session Log

### 2026-05-15 — תיעוד ראשוני + הפעלת הסקיל בפועל [shipped]
- **What was done:** הופעל `obsidian-vault-workflow` בפעם הראשונה בפרויקט. נוצרו 13 קבצי vault לפי התקן: `_index.md` ראשי, `Meeting Notes/_index.md`, ו-11 קבצי נושא. גלעד ביקש להחיל את הסקיל אוטומטית בתחילת כל סשן/פקודה.
- **Decisions:** קובץ אחד פר קובץ-פרויקט-שלנו (gal/irit/bar/barak/CLAUDE.md/env/...); קובץ מאחד לכל אוסף סקילים חיצוני (superpowers, obsidian). זה איזון בין דרישת המשתמש לתיעוד מפורט לבין דרישת הסקיל ל-one-file-per-topic.
- **Notes / Caveats:** ההחלטה להפעיל את הסקיל בכל סשן נשמרה ב-memory של Claude. ה-vault מוחרג ב-`.gitignore` נכון לעכשיו.
- **Related:** [[superpowers-plugin]], [[project-architecture]], [[root-gitignore]]

### 2026-05-15 — רענון: הסרת Open Question שנפתר [shipped]
- **What was done:** השאלה הפתוחה על `vault/` ב-`.gitignore` הוסרה — נפתרה כשהוסר `vault/` מההחרגה ב-commit `d4996ab`.
- **Decisions:** סריקת רענון לכל קבצי vault — שני קבצים נמצאו לא מסונכרנים עם המציאות הנוכחית.
- **Notes / Caveats:** ה-Session Log הקודם הזכיר שה-vault מוחרג; הערה זו כבר לא תקפה.
- **Related:** [[root-gitignore]]
