import sys
import os


sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import unittest
import datetime
import numpy as np
import gurobipy as gp

from src import LPSolver, ExcelTool
from util import Person, Schedule, Shift


class TestLPSolver(unittest.TestCase):
    def setUp(self):
        np.random.seed(42)  # For reproducible results
        self.schedule = self.create_four_week_schedule()
        self.people_config = self.create_people_config()
        self.add_people_to_schedule()
        self.schedule.calculate_non_sunday_hours()
        self.schedule.calculate_availability()
        self.max_hours = 120  # Effectively no limit for this test
        self.sunday_quota = 8
        self.solver = LPSolver(self.schedule, self.max_hours, self.sunday_quota)
        self.solver.setup_variables()
        self.solver.set_objective()
        self.solver.apply_constraints()

    def create_people_config(self):
        names = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']
        availability_percentages = np.linspace(0.1, 0.5, 15)  # 15 values from 0.1 to 0.4
        return [{'name': name, 'availability_percent': avail} for name, avail in zip(names, availability_percentages)]

    def create_four_week_schedule(self):
        schedule = Schedule()
        start_date = datetime.date(2024, 1, 13)  # Starting on January 13th, 2024 (Monday)
        days_en_to_nl = {
            "Monday": "Maandag",
            "Tuesday": "Dinsdag",
            "Wednesday": "Woensdag",
            "Thursday": "Donderdag",
            "Friday": "Vrijdag",
            "Saturday": "Zaterdag",
            "Sunday": "Zondag"
        }
        for day_offset in range(28):  # 4 weeks * 7 days = 28 days
            current_date = start_date + datetime.timedelta(days=day_offset)
            day_en = current_date.strftime("%A")
            day_nl = days_en_to_nl[day_en]
            if day_en in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']:
                # Weekday shifts: Ochtend (5h), Middag (5h), Avond (6h)
                shifts = [
                    Shift(time="08:00-13:00", hours=5, persons_required=1, shift_type="Ochtend", day=day_nl, date=current_date),
                    Shift(time="13:00-18:00", hours=5, persons_required=1, shift_type="Middag", day=day_nl, date=current_date),
                    Shift(time="18:00-24:00", hours=6, persons_required=1, shift_type="Avond", day=day_nl, date=current_date),
                ]
            else:
                # Weekend shifts: Ochtend (8h), Avond (8h)
                shifts = [
                    Shift(time="08:00-16:00", hours=8, persons_required=1, shift_type="Ochtend", day=day_nl, date=current_date),
                    Shift(time="16:00-24:00", hours=8, persons_required=1, shift_type="Avond", day=day_nl, date=current_date),
                ]
            schedule.shifts.extend(shifts)
        return schedule

    def add_people_to_schedule(self):
        for config in self.people_config:
            person = Person(config['name'])
            total_shifts = len(self.schedule.shifts)
            availability = self.generate_availability(total_shifts, config['availability_percent'])
            person.availability = availability
            self.schedule.people[person.name] = person

    def generate_availability(self, total_shifts, availability_percent):
        num_available = int(availability_percent * total_shifts)
        availability = [0] * total_shifts
        indices = np.random.choice(total_shifts, num_available, replace=False)
        for idx in indices:
            availability[idx] = 1
        return availability

    def test_solver_assignment(self):
        self.solver.solve()
        # self.assertEqual(self.solver.model.status, gp.GRB.OPTIMAL)
        self.perform_constraint_checks()
        metrics = self.calculate_metrics()
        print("\nSolver Evaluation Metrics:")
        for key, value in metrics.items():
            print(f"{key}: {value}")

        ExcelTool.write_schedule(self.schedule, self.solver, "test_case_schedule.xlsx")
        ExcelTool.write_metrics(self.schedule, self.solver, "test_case_metrics.xlsx")


    def perform_constraint_checks(self):
        self.check_shift_assignments()
        self.check_availability()
        self.check_no_evening_to_morning()
        self.check_sunday_quota()
        self.check_max_one_sunday_shift()
        self.check_max_hours()

    def check_shift_assignments(self):
        for shift_idx, shift in enumerate(self.schedule.shifts):
            assigned = sum(int(self.solver.A[(name, shift_idx)].X > 0.5) for name in self.schedule.people)
            slack = self.solver.slack[shift_idx].X
            self.assertEqual(assigned + slack, shift.persons_required, f"Shift {shift_idx} assignment mismatch")

    def check_availability(self):
        for shift_idx, shift in enumerate(self.schedule.shifts):
            for name in self.schedule.people:
                if self.solver.A[(name, shift_idx)].X > 0.5:
                    self.assertEqual(self.schedule.people[name].availability[shift_idx], 1, f"{name} assigned to unavailable shift {shift_idx}")

    def check_no_evening_to_morning(self):
        for name in self.schedule.people:
            assigned_shifts = [i for i, shift in enumerate(self.schedule.shifts) if self.solver.A[(name, i)].X > 0.5]
            for i in assigned_shifts:
                if self.schedule.shifts[i].shift_type == 'Avond' and i + 1 < len(self.schedule.shifts):
                    next_shift = self.schedule.shifts[i + 1]
                    if next_shift.shift_type == 'Ochtend' and (i + 1) in assigned_shifts:
                        self.fail(f"{name} has Avond followed by Ochtend at shifts {i} and {i+1}")

    def check_sunday_quota(self):
        for name, person in self.schedule.people.items():
            if person.non_sunday_hours < self.sunday_quota:
                for shift_idx, shift in enumerate(self.schedule.shifts):
                    if shift.day == 'Zondag':
                        self.assertLessEqual(self.solver.A[(name, shift_idx)].X, 0.5, f"{name} with low quota assigned to Sunday shift {shift_idx}")

    def check_max_one_sunday_shift(self):
        for name in self.schedule.people:
            sunday_shifts = sum(self.solver.A[(name, i)].X > 0.5 for i, shift in enumerate(self.schedule.shifts) if shift.day == 'Zondag')
            self.assertLessEqual(sunday_shifts, 1, f"{name} has {sunday_shifts} Sunday shifts")

    def check_max_hours(self):
        for name, person in self.schedule.people.items():
            total = sum(self.solver.A[(name, i)].X * shift.hours for i, shift in enumerate(self.schedule.shifts))
            self.assertLessEqual(total, self.max_hours, f"{name} worked {total} hours")

    def calculate_metrics(self):
        metrics = {}
        total_shifts = len(self.schedule.shifts)
        total_slack = sum(slack.X for slack in self.solver.slack.values())
        filled = sum(shift.persons_required - slack.X for shift, slack in zip(self.schedule.shifts, self.solver.slack.values()))
        metrics['Filled Shifts (%)'] = (filled / (total_shifts * 2)) * 100  # 2 persons per shift

        total_regular_error = 0
        total_bonus_error = 0
        for name, person in self.schedule.people.items():
            regular_hours = sum(self.solver.A[(name, i)].X * shift.hours for i, shift in enumerate(self.schedule.shifts))
            bonus_hours = sum(self.solver.A[(name, i)].X * shift.bonus_hours for i, shift in enumerate(self.schedule.shifts))
            expected_regular = (person.available_regular_hours / self.schedule.get_total_available_regular()) * sum(shift.hours * shift.persons_required for shift in self.schedule.shifts)
            expected_bonus = (person.available_regular_hours / self.schedule.get_total_available_regular()) * sum(shift.bonus_hours * shift.persons_required for shift in self.schedule.shifts)
            total_regular_error += (expected_regular - regular_hours) ** 2
            total_bonus_error += (expected_bonus - bonus_hours) ** 2
            metrics[f'{name} Regular Hours'] = regular_hours
            metrics[f'{name} Bonus Hours'] = bonus_hours

            # Calculate share received and expected share
            received_regular_share = regular_hours / person.available_regular_hours if person.available_regular_hours > 0 else 0
            expected_regular_share  = expected_regular / person.available_regular_hours if person.available_regular_hours > 0 else 0
            received_bonus_share = bonus_hours / person.available_bonus_hours if person.available_bonus_hours > 0 else 0
            expected_bonus_share = expected_bonus / person.available_bonus_hours if person.available_bonus_hours > 0 else 0
            
            received_total_share = (regular_hours + bonus_hours) / person.total_available_hours if person.total_available_hours > 0 else 0
            expected_total_share = (expected_regular + expected_bonus) / person.total_available_hours if person.total_available_hours > 0 else 0

            # Print employee-specific metrics
            print(f"\nEmployee: {name}")
            print(f"Available Regular Hours: {person.available_regular_hours}")
            print(f"Available Bonus Hours: {person.available_bonus_hours}")
            print(f"Available Total Hours: {person.total_available_hours}")
            print(f"Regular Hours Assigned: {regular_hours}")
            print(f"Bonus Hours Assigned: {bonus_hours}")
            print(f"Share Received Regular Hours: {received_regular_share:.5f}")
            print(f"Expected Share Regular Hours: {expected_regular_share:.5f}")
            print(f"Share Received Bonus Hours: {received_bonus_share:.5f}")
            print(f"Expected Share Bonus Hours:: {expected_bonus_share:.5f}")
            print(f"Share Received Total Hours: {received_total_share:.5f}")
            print(f"Expected Share Total Hours: {expected_total_share:.5f}")

        metrics['Regular Hours Fairness (SSE)'] = total_regular_error
        metrics['Bonus Hours Fairness (SSE)'] = total_bonus_error
        return metrics

if __name__ == '__main__':
    unittest.main()