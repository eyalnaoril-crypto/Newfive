# Meeting Notes — Index

תיעוד טכני של כל קובץ בפרויקט: מה הוא עושה, למי הוא שייך, ולאילו קבצים אחרים הוא קשור.

## נושאים

### מבט-על
- [[project-architecture]] — הצוות הדיגיטלי, זרימת העבודה, ומבנה התיקיות
- [[eyal-ceo-agent-prd]] — PRD ליצירת סוכן אייל כסוכן עצמאי (טרם בוצע — `[planned]`)

### קבצי שורש (Root Config)
- [[root-claude-md]] — `CLAUDE.md` — פרסונת אייל המנכ"ל
- [[root-readme]] — `README.md` — מסמך הסבר ראשוני (ירושה מפרויקט קודם, דורש עדכון)
- [[root-env-config]] — `.env` + `.env.example` — קונפיגורציה ארגונית וספי חריגים
- [[root-gitignore]] — `.gitignore` — קבצים שלא עולים ל-git
- [[root-package-json]] — `package.json` — תלויות Node.js
- [[claude-settings-local]] — `.claude/settings.local.json` — הגדרות מקומיות של Claude Code

### סוכנים (כל סוכן + הסקיל שלו)
- [[agent-gal]] — **גל** — ניתוח Excel, KPI, גרפים
- [[agent-irit]] — **אירית** — מצגות PowerPoint להנהלה
- [[agent-bar]] — **בר** — דוחות Word
- [[agent-barak]] — **ברק** — הפקת PDF להפצה

### סקילים חיצוניים
- [[superpowers-plugin]] — 14 סקילי `obra/superpowers` (תהליכי פיתוח, debugging, code review)
- [[obsidian-skills]] — 3 סקילי Obsidian (כולל obsidian-vault-workflow שמנהל את ה-vault הזה)
