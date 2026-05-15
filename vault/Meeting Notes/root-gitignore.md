# .gitignore — קבצים שלא עולים ל-git

## Overview

מגדיר אילו קבצים ותיקיות לא יעלו למאגר GitHub. נוצר בסשן ה-`git init` הראשון, מטרתו העיקרית: **להגן על `.env` שמכיל ערכים רגישים** (מפתחות API, ספי חריגים פנימיים), ולמנוע העלאת תוצרים גדולים (output/) או חומר פרטי (references/, vault/).

**שייכות:** קונפיגורציה כלל-מערכתית, חיונית לאבטחת המידע.

## Open Questions

- האם להוסיף `*.tmp`, `*.bak`, `__pycache__/` עבור עבודה עם Python בעיבוד מקדים?
- האם להוסיף הגנה ספציפית על קובצי `.docx`/`.xlsx` שמכילים נתוני נהגים פרטיים?

## תוכן נוכחי

```
node_modules/
.env
.env.local
output/
references/
.claude/settings.local.json
.DS_Store
*.log
```

## הצדקות לכל החרגה

| תיקייה/קובץ | סיבה |
|--------------|------|
| `node_modules/` | תלויות Node — נוצרים מ-`npm install` |
| `.env` / `.env.local` | מפתחות API וערכים פנימיים |
| `output/` | תוצרי הסוכנים — גדולים, מתחדשים, אישיים |
| `references/` | חומרי מותג ויזואליים פרטיים |
| `.claude/settings.local.json` | הגדרות Claude Code פרטיות לכל משתמש |
| `.DS_Store` / `*.log` | קבצים סטנדרטיים שלא רלוונטיים ל-repo |

## קבצים קשורים

- **מגן על:** [[root-env-config]] (`.env` שלא יעלה ל-git).
- **משפיע על תיעוד זה:** ה-vault עצמו מוחרג כרגע — לשקול שינוי.

## Session Log

### 2026-05-15 — תיעוד ראשוני + פלאג vault/ [shipped]
- **What was done:** מיפוי תוכן הקובץ והסיבות לכל החרגה. סומן פלאג ש-`vault/` עשוי לדרוש שינוי החלטה.
- **Decisions:** הקובץ נוצר בסשן `git init` המקורי. כל ההחרגות אושרו על ידי גלעד.
- **Notes / Caveats:** אם המשתמש ירצה לחלוק את ה-vault דרך GitHub — צריך להסיר את השורה `vault/` ולעשות commit חדש.
- **Related:** [[root-env-config]], [[project-architecture]]

### 2026-05-15 — הסרת vault/ מההחרגה [shipped]
- **What was done:** הוסרה השורה `vault/` מ-`.gitignore`. תיעוד ה-vault יעלה כעת ל-GitHub עם כל commit.
- **Decisions:** ה-vault הוא תיעוד פרויקטי שיתופי (לא זיכרון פרטי) — שווה לחלוק כדי שתיעוד מקיף יהיה זמין בכל מכשיר ולכל בעלי גישה ל-repo.
- **Notes / Caveats:** `.env`, `output/`, `references/`, ו-`.claude/settings.local.json` נשארים מוחרגים.
- **Related:** [[project-architecture]], [[obsidian-skills]]
