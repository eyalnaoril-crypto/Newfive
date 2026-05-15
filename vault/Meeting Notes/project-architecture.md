# Project Architecture — ארכיטקטורת המערכת

## Overview

**מערכת לניתוח צי רכב** של G1-Group. הפלטפורמה מקבלת קבצי Excel של צי בן ~600 כלי רכב המאוגדים בהיררכיית **קבוצה → חטיבה → סניף**, ומפיקה תוצרים ניהוליים מוכנים: Excel, PowerPoint, Word ו-PDF.

המבנה: **צוות סוכנים בראשות מנכ"ל דיגיטלי (אייל)**. אייל מקבל בקשות מגלעד (מנהל הצי האנושי), מקצה משימות לסוכן המתאים, מאחד תוצרים, ומחזיר סיכום ניהולי. הצוות עשוי להתרחב בעתיד.

## Open Questions

- האם להוסיף סוכנים נוספים (לדוגמה — סוכן Drill-down אינטראקטיבי, סוכן Telemetry חי) בשלבים מאוחרים יותר של ה-roadmap (§15 ב-PRD)?
- האם נדרש מסך UI פיזי (Dashboard), או שכל האינטראקציה דרך Claude Code בלבד?
- האם תהיה אינטגרציה ל-Linear/Notion/Slack לתיוג תוצרים?

## הצוות

| סוכן | תפקיד | סקיל ייעודי | פלט |
|------|--------|--------------|-----|
| **אייל** | מנכ"ל דיגיטלי, מתאם | פרסונה ב-[[root-claude-md|CLAUDE.md]] | סיכומים ניהוליים |
| **גל** | אנליסט Excel | `gal-excel-fleet-analyst` | `.xlsx` עם KPI וגרפים |
| **אירית** | מצגות הנהלה | `irit-fleet-presentation` | `.pptx` RTL |
| **בר** | כותב דוחות | `bar-fleet-word-report` | `.docx` Tahoma 11 |
| **ברק** | מפיק PDF | `barak-fleet-pdf-publisher` | `.pdf` חתום וגרסה |

## זרימת עבודה טיפוסית

1. גלעד פונה לאייל עם בקשה ומעלה קבצי Excel.
2. אייל מבין את ההקשר, בוחר רמת חתך (סניף/חטיבה/קבוצה) וסוג תוצר.
3. **גל מופעל ראשון** — בונה את שכבת הנתונים והאמת (KPI + גיליונות מסכמים).
4. אייל סוקר את ממצאי גל, מזהה חריגים, מחליט אילו סוכנים נוספים להפעיל.
5. אירית / בר / ברק רצים (במקביל כשאפשר) על ממצאים מאושרים.
6. אייל מאחד תוצרים, מחזיר לגלעד סיכום + נתיבי קבצים.

## מבנה תיקיות

```
Newfive/
├── CLAUDE.md                   → [[root-claude-md]]
├── README.md                   → [[root-readme]]
├── .env / .env.example         → [[root-env-config]]
├── .gitignore                  → [[root-gitignore]]
├── package.json                → [[root-package-json]]
├── .claude/
│   ├── settings.local.json     → [[claude-settings-local]]
│   ├── agents/                 → סוכני המשימות (gal/irit/bar/barak)
│   ├── skills/                 → סקילים — ידע תפעולי לכל סוכן + סקילים חיצוניים
│   └── commands/               → סלאש-קומנדות (טרם נוצרו)
├── output/                     → תוצרי הסוכנים (gitignored)
├── references/                 → חומרי מותג (gitignored)
└── vault/                      → תיעוד ארוך-טווח (הקבצים האלה)
```

## מקורות מידע

- **PRD מלא:** `OneDrive/Claude/רכב/PRD עבור פרוייקט ניתוח למחלקת הרכב1.docx`
- **GitHub:** https://github.com/eyalnaoril-crypto/Newfive

## Session Log

### 2026-05-15 — הקמת תיעוד ראשוני של הפרויקט [shipped]
- **What was done:** נוצר vault ראשון, מפת ארכיטקטורה מסכמת, וקבצי תיעוד לכל קובץ פרויקט.
- **Decisions:** קובץ פר-קובץ-קוד לקבצים שלנו; קובץ מאחד לכל אוסף סקילים חיצוני (superpowers/obsidian).
- **Notes / Caveats:** ה-PRD גרסה ראשונית — שלבי roadmap 2-4 (§15) טרם תוכננו במפורט.
- **Related:** [[root-claude-md]], [[agent-gal]], [[agent-irit]], [[agent-bar]], [[agent-barak]], [[superpowers-plugin]]
