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
input/*
!input/.gitkeep
output/
references/
.claude/settings.local.json
.obsidian/
.DS_Store
*.log
```

## הצדקות לכל החרגה

| תיקייה/קובץ | סיבה |
|--------------|------|
| `node_modules/` | תלויות Node — נוצרים מ-`npm install` |
| `.env` / `.env.local` | מפתחות API וערכים פנימיים |
| `input/*` (+ `!input/.gitkeep`) | נתוני צי גולמיים — מכילים מספרי רישוי, עלויות, פרטי נהגים. ה-`.gitkeep` שומר את התיקייה ב-repo כדי שגל ידע איפה לחפש |
| `output/` | תוצרי הסוכנים — גדולים, מתחדשים, אישיים |
| `references/` | חומרי מותג ויזואליים פרטיים |
| `.claude/settings.local.json` | הגדרות Claude Code פרטיות לכל משתמש |
| `.obsidian/` | תיקיית הגדרות מקומיות של אפליקציית Obsidian (workspace state, plugins מקומיים) |
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

### 2026-05-15 — הוספת .obsidian/ להחרגה [shipped]
- **What was done:** נוספה השורה `.obsidian/` ל-`.gitignore` — תיקיית הגדרות workspace של אפליקציית Obsidian נשמרת מקומית.
- **Decisions:** ההגדרות האלה (workspace layout, hot reload, plugins מקומיים) שונות ממכשיר למכשיר ולא רלוונטיות לתיעוד הפרויקטי עצמו ב-vault.
- **Notes / Caveats:** אם בעתיד נרצה להאחד הגדרות תצוגה בין מכשירים (לדוגמה קונפיגורציית graph view של ה-vault) — נשקול שוב.
- **Related:** [[obsidian-skills]]

### 2026-05-15 — החרגת תוכן input/ + שמירת התיקייה דרך .gitkeep [shipped]
- **What was done:** נוסף `input/*` + `!input/.gitkeep` ל-`.gitignore`. קובץ ריק `input/.gitkeep` נוצר כדי לשמור את מבנה התיקייה ב-repo. נתוני התאונות (`תאונות רבעון 1-2026.xls`) נשארים מקומיים בלבד.
- **Decisions:** ה-repo ציבורי ב-GitHub, ונתוני צי כוללים מספרי רישוי, עלויות ופרטי נהגים — סיכון פרטיות מוחשי. הפתרון הסטנדרטי `data-dir/*` + `.gitkeep` משאיר את הפרטיות בידי גלעד תוך שמירה על מבנה הפרויקט גלוי.
- **Notes / Caveats:** [[agent-gal]] עדיין סורק את `input/` באמצעות Glob — הסריקה תעבוד מקומית בלי שינוי, אבל קלון רענן ב-GitHub יקבל תיקייה ריקה (וגלעד יצטרך להעלות קבצים בעצמו).
- **Related:** [[agent-gal]], [[root-env-config]]
