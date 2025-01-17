from src.excel_writer import ExcelTool
from src.lp_solver import LPSolver

### CONFIGURATION ###
make_library = True
max_hours = 100
sunday_quota = 20

# Read schedule from Excel
schedule = ExcelTool.read_availability("Beschikbaarheid_Mock_Full.xlsx")
schedule.calculate_non_sunday_hours()

solver = LPSolver(schedule, max_hours, sunday_quota)
solver.setup_variables()
solver.set_objective()
solver.apply_constraints()
solver.solve()

# Save schedule to file
ExcelTool.write_schedule(schedule, solver, "Final_Schedule.xlsx")
