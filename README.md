# Roster Optimizer

This project builds a monthly **rooster** based on an availability Excel file and a mixed-integer model. 

---

## How it works

1. **Input (Beschikbaarheid.xlsx)**  
   - Reads the `"Hele Team"` sheet from the availability file.  
   - Parses:
     - Shifts: day, date, type (Ochtend/Middag/Avond), start–end time, hours, persons required.
     - People: binary availability per shift, total/non-Sunday availability, etc.
   - Builds an internal `Schedule` with `Shift` and `Person` objects.

2. **Optimization (Gurobi model)**  
   Implemented in `lp_solver.py`:

   - Binary decision variables:  
     - `A[person, shift] = 1` if a person works a shift, 0 otherwise.
   - **Fairness objective**:  
     - Minimizes squared error between each person’s *expected* share of hours (based on availability) and their *assigned* hours.  
     - Supports separate weighting for regular vs. bonus hours.
   - **Shift coverage**:  
     - Each shift requires a fixed number of people.  
     - If this is impossible, a **slack variable** fills the gap with a very high penalty → unfilled / “onhaalbaar” shifts.
   - **Constraints include**:
     - Respect availability (no assignment if unavailable).
     - No evening → next-morning combos.
     - No long day chains (e.g. morning + afternoon + evening).
     - Sunday rules:
       - Only people above a minimum non-Sunday (20 hours default) quota can work on Sunday.
       - Max 1 Sunday shift per person.
     - Max total hours per person (`max_hours`).

3. **Output (Excel roosters)**  
   - `excel_writer.py` writes the solution back to an `.xlsx` file per poule.  
   - Shifts are grouped by day and show the assigned names.  
   - Unfilled positions (positive slack) are marked as “no one available” / “onhaalbaar” in the sheet.

## Usage

Carefully look at the layout for `Beschikbaarheid_Mock_Full.xlsx`, this same structure must be adhered to. See the `ExcelTool` for how avaliability is parsed from the excel sheet. 

1. **Install dependencies**

```bash
   pip install -r requirements.txt
```

Make sure Gurobi is installed and licensed. You can get a license via the TU

2. **Prepare input**
   * Place your availability file (e.g. `Beschikbaarheid.xlsx` / `Beschikbaarheid_Mock_Full.xlsx`) in the repo.
   * Ensure the sheet and column structure match the expected format see the examples attached.

3. **Configure**
   * Open `main.py` and adjust:
     * Path to the availability file.
     * Parameters like `max_hours`, `sunday_quota`, etc.

4. **Run**
```bash
   python main.py
```

   The script will:

   * Parse the availability.
   * Build and solve the Gurobi model.
   * Write roster file 

5. **Verify (optional)**

   * `verification.py` and `test_case_bonus.py` contain checks/test cases to validate that constraints and bonus-hour behavior work as expected.
