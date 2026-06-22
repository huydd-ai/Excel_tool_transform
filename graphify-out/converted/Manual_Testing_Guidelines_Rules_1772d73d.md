<!-- converted from Manual_Testing_Guidelines_Rules.xlsx -->

## Sheet: General Guidelines
| Category | Requirement / Rule | Description |
| --- | --- | --- |
| Tester Role | Manual Tester | Responsible for end-to-end verification of game logic and UI based on GD docs. |
| Bug Reporting | Standard Template | Must include: Feature ID, Scenario, Input/State, Expected Result, Actual Result, Severity (Low-Critical). |
| Severity Definition | Critical | Blocking progress, data loss, or economic exploitation (e.g., unlimited gold/hearts). |
| Severity Definition | High | Functional failure in a major feature (e.g., Royal Pass rewards not unlocking). |
| Severity Definition | Medium | UI/UX issues that affect experience but not progression. |
| Network Testing | Connectivity | Always test features under 'Airplane Mode' and 'Fluctuating Network' to verify sync logic. |
## Sheet: Feature Rules
| Feature | Logic Item | Rule / Condition | Expected Behavior |
| --- | --- | --- | --- |
| Heart System | Passive Regen | Count < Max (5) | 1 Heart added every 20 mins. Timer persists offline. |
| Heart System | Consumption | Fail / Exit / Restart | Deduct 1 Heart ONLY after confirming warning popup. |
| Heart System | Level Win | Win Level | NO Heart deducted. |
| Heart System | Unlimited Heart | Active Duration > 0 | Deductions = 0. Standard timer hidden; shows expiry countdown. |
| Royal Pass | EXP Progression | Win 1 Level | Grant +1 EXP exactly. No EXP on loss. |
| Royal Pass | Max Lives Buff | Gold/Ultimate Pass | Max hearts increase from 5 to 8 during season active duration. |
| Royal Pass | Tier Unlocks | EXP Threshold reached | Instantly unlock rewards. Ultimate Pass unlocks +2 tiers ahead. |
| Lava Quest | Streak Logic | 7 Level Wins | Must be consecutive. 1 loss/quit/restart = 0 streak + 10 min cooldown. |
| Lava Quest | Reward Pool | Shared Payout | 5000 Coins / (Final BOTs + 1 user). BOTs eliminated per level. |
## Sheet: Edge Cases
| ID | Scenario | Condition | Required Handling (Recovery) |
| --- | --- | --- | --- |
| HEART-EDG-01 | Clock Manipulation | User sets phone clock back | Local UI may show timer, but Server/Secure time must override and correct. |
| HEART-EDG-02 | App Kill Mid-level | Force close game | Deduct 1 Heart on relaunch if level was active and not finished. |
| RPS-EDG-03 | Simultaneous Claim | Same account on 2 devices | Idempotency check: First claim succeeds, second returns 'already claimed'. |
| LVQ-EDG-01 | Event Expired in Level | 24h ends while playing | If win, allow advance/reward. Show 'Event Over' after level ends. |
| LVQ-EDG-02 | Bypass Cooldown | Change local clock | Server-side validation must block entry until 10-minute real-time passed. |
## Sheet: Release Checklist
| Check ID | Verify Item | Pass/Fail | Notes |
| --- | --- | --- | --- |
| CHK-01 | Hearts not deducted on Win. |  |  |
| CHK-02 | Gold Pass updates Heart Max to 8. |  |  |
| CHK-03 | Lava Quest streak resets on manual restart. |  |  |
| CHK-04 | Timer format is correct (HH:MM:SS). |  |  |
| CHK-05 | Buy +5 Hearts (150 Money) works offline/online sync. |  |  |
| CHK-06 | Royal Pass EXP updates in UI instantly after level win. |  |  |
| CHK-07 | Cooldown screen visible for 10 mins after Lava Quest loss. |  |  |