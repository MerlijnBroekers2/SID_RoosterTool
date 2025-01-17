import sys
import os

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import pandas as pd
import unittest
from src import LPSolver
from src import ExcelTool

class TestSchedule(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # Load schedule from test file
        cls.schedule = ExcelTool.read_availability("Beschikbaarheid_Mock_Full.xlsx")
        cls.schedule.calculate_non_sunday_hours()
        cls.solver = LPSolver(cls.schedule, max_hours=100, sunday_quota=20)
        cls.solver.setup_variables()
        cls.solver.set_objective()
        cls.solver.apply_constraints()
        cls.solver.solve()
    
    def test_total_hours_per_person(self):
        """Verify that total assigned hours per person match the constraints."""
        total_hours = {person: 0 for person in self.schedule.people}
        
        for shift_idx, shift in enumerate(self.schedule.shifts):
            for person_name in self.schedule.people:
                if self.solver.A[(person_name, shift_idx)].X > 0.5:
                    total_hours[person_name] += shift.hours
        
        # Convert to DataFrame for better visualization
        df = pd.DataFrame(total_hours.items(), columns=["Person", "Total Hours Assigned"])
        print("\nAssigned Hours Per Person:\n", df.to_string(index=False))
        
        # Check if no one exceeds max_hours constraint
        for person_name, hours in total_hours.items():
            self.assertLessEqual(hours, self.solver.max_hours, f"{person_name} exceeds max hours limit!")

if __name__ == "__main__":
    unittest.main()
