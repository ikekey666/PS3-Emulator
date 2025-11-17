# SQUEEZE VBA MODULE - CLEAN SPECIFICATION
**Version:** 2.0 (Deduplicated & Consolidated)  
**Date:** 2025-11-17  
**Type:** Excel VBA - ThisWorkbook Module  
**UI Standard:** Option B (Light Neutral - User Friendly)  
**Compatibility:** Excel 2016+ (Windows/Mac)

---

## 1. CORE CONCEPT

**Almir Budalica Squeeze** is a single-module Excel VBA engine for:
- **Exhaustive k-sum combination building** (k=5 or k=6)
- **Instance-aware destructive mode** with cell-level tracking
- **Global deduplication** by value-set (order-agnostic)
- **Quality scoring** (QScore, Grade, Stars, Tier)
- **Smart output management** with optional tier filtering

### Key Principles
- One combination uses each cell **exactly once**
- Destructive mode clears cells and marks with red fill (fully undoable)
- Non-destructive variations preserve original pool
- All existing Squeeze_Sum_* sheets preloaded to avoid duplicate output
- No visible footer stats; hidden metadata stored in SQUEEZE_META (VeryHidden sheet)

---

## 2. POOL MODES & SELECTION

### Mode 1: Manual Multi-Area
- Load numbers directly from current selection
- Numeric values only (no text)
- Counts visible, non-hidden cells

### Mode 2: Auto-Detect K/N
- Select a grid/range
- Engine calculates:
  - **NAIVE K:** Count of all numeric columns
  - **CORRECT K:** Numeric columns minus 'Sum/Total' headers
- Banner displays: `[NAIVE K: x | CORRECT K: y]`

### Game Aliases (Auto-Mapping)
| Game | Key | Config | K | N |
|------|-----|--------|---|---|
| Powerball | PB | 5/69 | 5 | 69 |
| Mega Millions | MM | 5/70 | 5 | 70 |
| Lotto / Pick 6 | L649/P6 | 6/49 | 6 | 49 |
| Cash 5 | C5 | 5/43 | 5 | 43 |
| Treasure Hunt | TH | 5/30 | 5 | 30 |
| Match 6 | M6 | 6/49 | 6 | 49 |
| Custom | CUSTOM | User-defined | 5 or 6 | User |

---

## 3. TARGET SUM SYSTEM

- **0 (Zero):** Accept ANY sum (no constraint)
- **Positive Value:** Exact target sum required; only matching combinations generated
- **Smart Suggestion:** Engine calculates min/max possible sums from pool and suggests average
- **Sum Validation:** Only combinations that match target are output

---

## 4. BUILD TYPES & USER PROMPTS

Four standardized build types with configurable parameters:

### Option 1: MINIMAL
- Pool reduction: 50%
- Fewest constraints
- Lowest cost
- Best for: Budget-conscious users, beginners

### Option 2: SMART ⭐ *Recommended*
- Pool reduction: 70%
- Balanced constraints
- Moderate cost (~6–8 tickets)
- Best for: Balanced approach, most users

### Option 3: OPTIMAL
- Pool reduction: 85%
- Strong pattern-based constraints
- Higher initial cost, better odds
- Best for: Maximum optimization, serious players

### Option 4: BALANCED / UNBALANCED
- **BALANCED:** Even distribution across groups
- **UNBALANCED:** Weighted toward hot trends
- Cost varies by strategy

### Configurable Prompt Parameters
- Pool reduction: `[FULL_POOL]` → `[TARGET_SIZE]`
- Improvement multiplier: `[20x / 50x / 100x]`
- Combination budget: `[6 / 8 / 10 / Custom]`
- Wheeling integration: `[Yes / No]`
- Reduction method: `[Trend-based / Statistical / Hybrid]`
- Optimize for: `[Jackpot / Multi-tier / Balanced]`

---

## 5. HARD RULES & GUARDS

| Rule | Setting | Behavior |
|------|---------|----------|
| **Max Repeat Per Combo** | 0 | No duplicate values within a single group |
| **Global Dedup** | YES | Prevents same value-set in any order across passes |
| **Odd/Even Guard** | Enabled | Rejects all-odd or all-even groups |
| **Auto S_MAX (Span)** | 15 | Maximum range (max − min) in combination |
| **Auto R_MAX (Run)** | 2 | Longest chain of consecutive numbers allowed |
| **Pruned Search** | Enabled | Prefix/suffix sums bound feasible space |

---

## 6. QUALITY SCORING & GRADING

### Scoring Components

**GH Score (Gail Howard)** — 50% of composite
- Odd/Even balance (ideal 50/50): 25 pts
- High/Low balance (game-specific midpoint): 25 pts
- Consecutive run limit penalty: 25 pts
- Span reasonableness reward: 25 pts

**AB Score (Advantage Bias)** — 50% of composite
- Divisible by 5: +10 pts per number
- Divisible by 7: +10 pts per number
- 16–24 band: +5 pts per number
- Even numbers: +5 pts per number
- Base: 50 pts

**AG Score (Anti-Gap)**
- Gap consistency (ideal ~5 units)
- Penalizes irregular spacing
- Detects and penalizes tight clustering

**IS Score (Instance Score)**
- All unique values: 100%
- Duplicate penalty: Reduced proportionally

**Composite Score** = Average of GH + AB + AG + IS (0–100)

### Grading & Tier System

| Grade | Score | Tier | Color | Use |
|-------|-------|------|-------|-----|
| A+ | 90–100 | BEST | Deep Green (0,100,0) | Best combinations |
| A | 80–89 | GOOD | Green (0,150,0) | Good combinations |
| B | 70–79 | — | Light Green (150,200,150) | Modest combinations |
| C | 60–69 | MODEST | Pale Gold (255,230,180) | — |
| D | 50–59 | — | Orange (255,180,100) | Low combinations |
| F | 0–49 | LOW | Red (255,100,100) | Filtered by default |

**Tier Thresholds:**
- **BEST:** QScore ≥ 90
- **GOOD:** QScore ≥ 80
- **MODEST:** QScore ≥ 65
- **LOW:** QScore < 65 (excluded by default)

**Tier Filter at Build Time:** B=Best only, G=Good only, M=Modest only, blank=ALL (default)

---

## 7. DESTRUCTIVE MODE & POOL REDUCTION

### Destructive Execution
When destructive mode is **ON**:
1. Consumed numbers are **cleared in place**
2. Cell background set to **solid red fill**
3. Every consumed cell logged to **UNDO sheet** (Address, Value)
4. Cells marked with **✓ checkmark** or grayed font (RGB 150,150,150)
5. Numbers removed from active pool after use
6. **One-time use per cell:** Each cell used exactly once per session

### Clear-As-You-Go
- Default: **YES**
- Marks cells immediately after consumption
- Allows real-time pool visibility

### Exhaustion Mode
- Continues until no valid combinations remain
- Respects all guards and constraints
- Reports final pool status

---

## 8. VARIATIONS SYSTEM

After base pass completes, optionally generate additional combinations:

### Variation 1: Sum_Variation
- Type: Non-destructive
- Pool: Remaining cells after base pass
- Maintains cell references
- Limit options: 0=Skip, −1=Unlimited, N=Exact count

### Variation 2: Sum_Variation_Copy
- Type: Shuffled duplicate of base
- Rows marked with **'c' suffix**
- Random row arrangement
- Count matches base pass

### Variation 3: Sum_Variation_More
- Type: Exhaust leftovers
- Pool: All remaining unused cells
- Generates until pool exhausted
- Limit options: 0=Skip, −1=Unlimited, N=Exact count

### Variation 4: Global_Shuffle
- Type: Complete random rebuild
- Pool: Original snapshot (before destruction)
- User-defined number of passes (not fixed)
- Each pass: Independent random order
- Limit options: 0=None, 1+=That many passes

### AUTO Mode
- Type: Continuous sweeps
- Repeats until exhausted
- Options: 0=Cancel, −1=Until exhausted, N=Exact sweeps
- Naming: AUTO_Sweep_1, AUTO_Sweep_2, etc.

---

## 9. OUTPUT LAYOUT

### Column Structure
| # | Columns | Header | Description |
|---|---------|--------|-------------|
| 1 | Rank | Rank | Combination number |
| 2 | N1...Nk | Numbers | Comma-separated values |
| 3 | Sum | Sum | Total of combination |
| 4 | QScore | QScore | Composite quality score (0–100) |
| 5 | Grade | Grade | Letter grade (A+–F) |
| 6 | Stars | Stars | Visual rating (★★★ to ■■■) |
| 7 | Sqz | Sqz | Internal squeeze metric |
| 8 | Tier | Tier | BEST/GOOD/MODEST/LOW |

### Sheet Naming
**Same-Sheet Mode:**
- Output positioned one column to the right of farthest "Matched (sum)" / "Unmatched (sum)"
- Auto-created if missing
- No overlap with input pool

**Splitter Mode:**
- Creates: `Squeeze_Sum_<Target>_<YYYYMMDD>[_B|_G|_M]`
- Suffix appended if tier-filtered (B=Best, G=Good, M=Modest)

### Row Structure
- **Row 1:** Game banner: `<Game> [NAIVE K: x | CORRECT K: y]` (Mode 2 only); appends `(Tier: B/G/M)` when filtering
- **Row 2:** Headers as above
- **Row 3+:** Data rows

---

## 10. UNDO SYSTEM

### Undo History
- Tracks: `{Address, Value, Color}`
- Maintained throughout session
- Per-session only (cleared on workbook close)

### Undo Functions
- **Undo Last (Ctrl+Z):** Restore most recent combination
- **Undo All (Ctrl+Shift+Z):** Restore all with confirmation
- **BULK_RESTORE_ALL:** Repopulates values and removes red fill

---

## 11. QUALITY DASHBOARD (Ctrl+Shift+Q)

### Operation
1. Sweeps all `Squeeze_Sum_*` sheets
2. Global-dedups by value-set
3. Outputs consolidated table
4. Prompts:
   - Include LOW tier? (default: No)
   - Tier filter: blank=ALL, B, G, or M

### Sorting
1. Tier: BEST → GOOD → MODEST → LOW
2. QScore (descending)
3. Sum (descending)

### Metadata Storage
- Hidden stats in **SQUEEZE_META (VeryHidden sheet)**
- Per-run counters: Solutions, Nodes, Pruned(len/sum)
- Legacy "-- Stats:" lines auto-cleared from results

---

## 12. UI STANDARD: OPTION B (LIGHT NEUTRAL)

**Visual Design:**
- Background: Light gray card
- Text: Dark gray (high contrast)
- Prompts: Clear, sequential, user-friendly
- Status: Inline HUD (progress, pool count)
- No timestamps in output
- Custom prefixes/aliases auto-detected

**Prompt Flow:**
1. Pool selection method (1/2/3/4/7)
2. Range selection (if manual)
3. Game selection or custom entry
4. Target sum (or 0 for any)
5. Visible count popup (statistics)
6. Build type selection (1/2/3/4)
7. Variations? (Yes/No)
8. Variation type (if yes)
9. Limits per variation

**No special cell selection required** — standard multi-select cell range support.

---

## 13. HOTKEY ASSIGNMENTS

| Hotkey | Function | Description |
|--------|----------|-------------|
| Ctrl+Shift+B | BuildSum_Lottery | Main entry point |
| Ctrl+Z | UndoLastCombination | Undo last combination |
| Ctrl+Shift+Z | UndoAllCombinations | Undo all with confirmation |
| Ctrl+Shift+Q | QualityDashboard | Consolidated quality view |

---

## 14. INTEGRATION WITH EXISTING SCRIPTS

### Compatibility Requirements
✓ Works alongside other VBA modules in same workbook  
✓ Respects existing sheet structures  
✓ Preloads existing `Squeeze_Sum_*` sheets to prevent duplicates  
✓ Safe block placement (no interference with other code)  
✓ Instance-aware (same cell value in different cells = different instances)  
✓ Maintains original data integrity (destructive mode fully reversible)

### For Developers
- All functions named with `Squeeze_` prefix
- Uses explicit variable declarations (Option Explicit)
- Arrays: `mPool()`, `mAddrs()` for active pool
- Snapshot: `mOrigVals()`, `mOrigAddrs()` for variations
- No external dependencies (no Dictionary, ArrayList, or API calls)
- Cross-platform compatible (Windows & Mac Excel)

---

## 15. DEFAULT SETTINGS SUMMARY

| Setting | Default | Description |
|---------|---------|-------------|
| Run Mode | D (Destructive) | Clears cells + red fill, with UNDO |
| Distinct Values | YES | No duplicates within single group |
| Global Dedup | YES | Prevents re-listing same set in any order |
| Global Shuffle | YES | Reshuffles pool between cycles |
| Quality Tiers | YES | Labels BEST/GOOD/MODEST |
| Clear-As-You-Go | YES | Marks cells immediately |
| Cycles | 1 | Base pass (tune as needed) |
| Leftover Sweeps | 0 | Additional exhaustion passes |
| Exhaust Until Cannot | YES | Stop when pool cannot form k |
| Tier Filter | Blank (ALL) | B/G/M for single-tier build |
| Output Choice | 1 (Same Sheet) | 1=Same, 2=Splitter |

---

## 16. QUICK START WORKFLOW

1. **Backup** your workbook (best practice before destructive mode)
2. **Select** numbers (Mode 1/2) or grid (Mode 9)
3. **Press Ctrl+Shift+B** → Follow prompts
4. **Review** results:
   - Row 1: Game banner
   - Row 2: Headers
   - Row 3+: Combinations (no stats footer)
5. **Optional:** Press **Ctrl+Shift+Q** for consolidated Quality Dashboard
6. **Undo:** Press **Ctrl+Z** or run **BULK_RESTORE_ALL** to restore original pool

---

## 17. PERFORMANCE & TECHNICAL

### Optimization
- Screen updating disabled during processing
- Events disabled during processing
- Calculation set to manual
- DoEvents every 10 combinations for responsiveness
- Status bar progress updates
- QuickSort for efficient pool pruning

### Data Structures
- Arrays: `mPool()`, `mAddrs()`
- Snapshot: `mOrigVals()`, `mOrigAddrs()`
- Collections: Undo history
- Range objects: Pool and output areas
- Type safety: Explicit declarations

### Exclusions
- SQUEEZE_OUTPUT_AREA: Never in pool
- Hidden cells: Excluded
- Non-numeric values: Ignored
- Negative numbers: Not collected
- Zero values: Not collected
- Used cells: Marked and excluded from pool

---

## 18. NOTES & TROUBLESHOOTING

**No groups output?**
- Verify Target (0=ignore), k (5/6), pool size
- Check odd/even guard, S_MAX, R_MAX caps
- Ensure sufficient pool for k combinations

**Duplicates appearing?**
- Ensure Global Dedup = YES
- Verify prior result sheets use Showcase headers
- Check preloading logic recognizes existing sheets

**Performance slow on large pools?**
- Increase S_MAX or R_MAX to reduce search space
- Use Minimal or Smart build type first
- Consider tier filtering (B/G only) to reduce output

**Undo not working?**
- Confirm UNDO sheet exists and is accessible
- Check workbook hasn't been saved/closed (session-based)
- Run BULK_RESTORE_ALL for complete restoration

---

## 19. DOCUMENT METADATA

**Title:** Almir Budalica Squeeze — VBA Specification (Clean)  
**Module Type:** Excel VBA ThisWorkbook Module  
**Features:** Destructive pool reduction, Instance-aware, Variations, Quality scoring, Tier filtering, Global dedup  
**Compatibility:** Excel 2016+ (Windows/Mac)  
**UI Standard:** Option B (Light Neutral)  
**License:** Free to use and modify  
**Generated:** 2025-11-17 03:03 UTC

---

*End of Specification*