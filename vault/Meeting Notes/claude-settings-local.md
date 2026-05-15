# .claude/settings.local.json — הגדרות Claude Code מקומיות

## Overview

קובץ הגדרות **מקומיות לכל משתמש** של Claude Code. תפקידו העיקרי: רשימת הרשאות (`permissions.allow`) שמאפשרת הרצה אוטומטית של פקודות Bash/PowerShell מסוימות בלי לבקש אישור בכל פעם. רשימה זו מצטברת לאורך זמן ככל שגלעד מאשר פקודות חוזרות.

**שייכות:** קונפיגורציית כלי הפיתוח של המשתמש האנושי (גלעד). **לא של אף סוכן.** מוחרג ב-`.gitignore` (זה אישי ושונה ממכשיר למכשיר).

## Open Questions

- האם להחיל את סקיל `fewer-permission-prompts` כדי לאחד permissions ולשפר חוויה?
- האם להעביר חלק מההרשאות לרמת `settings.json` (שיתופי לכל הצוות) בעתיד?
- ה-permission הספציפי ל-`TMPDIR=/tmp/tmp.62mITsz9jo/...` הוא חד-פעמי וחסר ערך — לנקות.

## תוכן נוכחי (סוגי הרשאות)

- `Bash(node *)`, `Bash(git add *)`, `Bash(git commit *)`, `Bash(git push *)` — פקודות פיתוח רגילות.
- `WebSearch`, `WebFetch(domain:ai.google.dev)` — גישה לאינטרנט מצומצמת.
- `PowerShell(...)` — שתי פקודות ספציפיות לקריאת ה-PRD מ-OneDrive (Expand-Archive ו-python-docx).
- `Bash(cat /c/Users/elada/.claude/plugins/cache/...)` — שריד מפרויקט קודם, לא רלוונטי.

## קבצים קשורים

- **מוחרג ב:** [[root-gitignore]] (לא יעלה ל-git — אישי).
- **קשור לסקיל:** `fewer-permission-prompts` (יוכל לעזור לנקות ולאחד).
- **משמש את:** כל הסשנים של Claude Code על הפרויקט.

## Session Log

### 2026-05-15 — תיעוד ראשוני [shipped]
- **What was done:** מיפוי תוכן הקובץ והאחריות שלו.
- **Decisions:** לא לערוך — זה קובץ אישי שמתעדכן אוטומטית כשגלעד מאשר פקודות.
- **Notes / Caveats:** יש שאריות מפרויקט קודם (`elada`, `Yuval`). לא חוסם — רק רעש.
- **Related:** [[root-gitignore]], [[project-architecture]]

### 2026-05-15 — רענון: נוספה הרשאת `Bash(find vault *)` [shipped]
- **What was done:** ב-Refresh pass זוהה ש-permission חדש (`Bash(find vault *)`) נוסף לקובץ אחרי סשני העבודה על ה-vault. זה תקין — Claude Code שמר אותה אוטומטית כשגלעד אישר את הפקודה.
- **Decisions:** לא נדרשת פעולה — זו ההתנהגות הצפויה של settings.local.json.
- **Notes / Caveats:** הקובץ ימשיך לגדול כך עם הזמן ויתכן שיצטרך לעבור consolidation דרך `fewer-permission-prompts`.
- **Related:** [[root-gitignore]]
