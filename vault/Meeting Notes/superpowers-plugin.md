# Superpowers Plugin — 14 סקילי obra/superpowers

## Overview

**Superpowers** הוא תוסף קוד-פתוח של [obra/superpowers](https://github.com/obra/superpowers) המספק 14 סקילי תהליך-פיתוח. הותקן בפרויקט בהעתקה ידנית (לא דרך מערכת `/plugin` של Claude Code). הסקילים תומכים בעבודה איכותית: brainstorming לפני קוד, תכנון, ביצוע, debugging שיטתי, code review, ו-test-driven development.

**שייכות:** סקילים **כלל-פרויקטיים** — לא משויכים לסוכן ספציפי. אייל וכל ארבעת הסוכנים יכולים להשתמש בהם.

**מיקום:** `.claude/skills/<skill-name>/SKILL.md` (כל סקיל בתיקייה משלו).

## Open Questions

- האם להחיל את `using-superpowers` כסקיל אוטומטי בתחילת כל סשן (כדי לקבוע סדר עבודה אחיד)?
- חלק מהסקילים מתאימים יותר לפרויקטי קוד מסורתיים. האם להגדיר ai-readme שמסביר אילו מהם רלוונטיים לפרויקט הזה (שהוא בעיקר orchestration)?
- האם להוסיף Pre-commit hook שמפעיל את `verification-before-completion` אוטומטית?

## רשימת 14 הסקילים

| סקיל | מתי להפעיל |
|------|-------------|
| **brainstorming** | לפני התחלת עבודה יצירתית (תכנון פיצ'ר, רכיב חדש) |
| **writing-plans** | כשיש spec/requirements למשימה רב-שלבית, לפני נגיעה בקוד |
| **executing-plans** | כשיש תוכנית כתובה לבצע ב-session נפרד עם checkpoints |
| **writing-skills** | יצירה/עריכה/בדיקה של סקילים |
| **systematic-debugging** | בכל bug, test failure, או התנהגות לא צפויה |
| **test-driven-development** | לפני כתיבת קוד פיצ'ר/bugfix |
| **verification-before-completion** | לפני declaration שעבודה הושלמה |
| **using-git-worktrees** | כשעבודה דורשת בידוד מ-workspace הנוכחי |
| **finishing-a-development-branch** | כשמימוש הושלם וצריך להחליט על merge/PR/cleanup |
| **requesting-code-review** | בסוף משימות, לפני merge |
| **receiving-code-review** | כשמקבלים פידבק על קוד |
| **dispatching-parallel-agents** | כשיש 2+ משימות עצמאיות במקביל |
| **subagent-driven-development** | ביצוע תוכניות implementation עם משימות עצמאיות |
| **using-superpowers** | בתחילת שיחה — מסביר איך למצוא ולהשתמש בסקילים |

## דוגמת שימוש בפרויקט שלנו

- **לפני בניית סוכן חדש:** `brainstorming` → `writing-plans` → `executing-plans` → `verification-before-completion`.
- **שינוי לסקיל קיים של גל:** `writing-skills` → `requesting-code-review`.
- **באג בתוצר שאירית הפיקה:** `systematic-debugging`.

## קבצים קשורים

- **מתועד ב-`.gitignore`?** לא — הסקילים האלה כן עולים ל-git (גלעד ירצה אותם זמינים בכל מכשיר).
- **אינדקס סקילים אחר:** [[obsidian-skills]] (3 סקילי Obsidian נפרדים).
- **ארכיטקטורה:** [[project-architecture]].

## Session Log

### 2026-05-15 — התקנה ידנית [shipped]
- **What was done:** `git clone --depth 1` של `obra/superpowers` ל-temp, ואז `cp -n` של תיקיית `skills/*` לתוך `.claude/skills/`. אין commands או agents ב-repo המקור.
- **Decisions:** התקנה ידנית (לא דרך `/plugin`) כי מערכת ה-plugins לא זמינה אצל גלעד. שימוש ב-`-n` (no-clobber) הבטיח שלא ידרוס סקילים שלי.
- **Notes / Caveats:** 14 סקילים, 43 קבצים, 8,323 שורות. Commit `0f8cadb`.
- **Related:** [[obsidian-skills]], [[project-architecture]]
