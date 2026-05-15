# package.json — תלויות Node.js

## Overview

קובץ `package.json` בשורש הפרויקט. שוכתב ב-2026-05-15 — כעת מצהיר על שם הפרויקט (`fleet-analysis-system`), גרסה `0.1.0`, ותיאור מערכת multi-agent של G1-Group. אין dependencies כרגע — הפרויקט רץ בתוך Claude Code, אין צורך ב-Node runtime.

**שייכות:** קובץ מטא-פרויקט (Node ecosystem). לא משויך לסוכן ספציפי. אם בעתיד נוסיף Node scripts לעיבוד מקדים של Excel — נוסיף dependencies כאן.

## Open Questions

- האם בעתיד נוסיף Node scripts לעיבוד מקדים של Excel (למשל המרת CSV לחתכים)? אם כן — להוסיף `pandas-equivalent`/`xlsx` dependencies.
- האם להוסיף `engines.node` כדי לקבע גרסת Node לסביבת הפיתוח?

## תוכן נוכחי

```json
{
  "name": "fleet-analysis-system",
  "version": "0.1.0",
  "description": "Fleet analysis multi-agent system for G1-Group — Excel/PowerPoint/Word/PDF outputs for ~600-vehicle fleet across group→division→branch hierarchy",
  "private": true,
  "type": "module"
}
```

## קבצים קשורים

- **קונפיגורציה:** [[root-env-config]] (`dotenv` עשוי לקרוא את `.env` אם יוסיפו Node scripts).
- **תיעוד אורך:** [[root-readme]] (גם הוא ירושה מפרויקט קודם).

## Session Log

### 2026-05-15 — תיעוד ראשוני + סימון לשכתוב [wip]
- **What was done:** זוהה שה-`package.json` ירושה מפרויקט קודם. תלויות לא בשימוש בפועל.
- **Decisions:** לא נמחק/נדרס בסשן הזה — מחכה לאישור מפורש.
- **Notes / Caveats:** ה-script `generate` יכשל כי `scripts/generate.mjs` לא קיים.
- **Related:** [[root-readme]], [[project-architecture]]

### 2026-05-15 — שכתוב מלא ל-fleet-analysis-system [shipped]
- **What was done:** דריסה מלאה — שם, גרסה ותיאור עודכנו. הוסרו `@google/genai` ו-`dotenv` (לא בשימוש) וה-script `generate` שהצביע לקובץ לא קיים.
- **Decisions:** השארנו `package.json` (לא מחקנו) כי הוא חסר עלות וייתכן צורך עתידי. `private: true` מונע פרסום בטעות ל-npm.
- **Notes / Caveats:** `dotenv` הוסר — אם בעתיד יהיה Node script שיקרא את `.env`, נחזיר אותו.
- **Related:** [[root-readme]], [[root-env-config]], [[project-architecture]]
