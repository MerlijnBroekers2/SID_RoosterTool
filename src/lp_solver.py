import gurobipy as gp
import numpy as np

class LPSolver:
    def __init__(self, schedule, max_hours, sunday_quota):
        self.schedule = schedule
        self.max_hours = max_hours
        self.sunday_quota = sunday_quota
        self.model = gp.Model()
        self.model.setParam('TimeLimit', 30)
        self.A = {}  # Decision variables
        self.slack = {}  # Slack variables for unfilled shifts

    def setup_variables(self):
        for shift_idx, shift in enumerate(self.schedule.shifts):
            for person_name, person in self.schedule.people.items():
                self.A[(person_name, shift_idx)] = self.model.addVar(vtype=gp.GRB.BINARY, name=f"A_{person_name}_{shift_idx}")

            # Add slack variable to represent unfilled shifts (if available people are fewer than required)
            self.slack[shift_idx] = self.model.addVar(vtype=gp.GRB.INTEGER, lb=0, name=f"Slack_{shift_idx}")

    def set_objective(self):
        error = {}
        objective = 0
        total_shift_hours = sum(shift.hours for shift in self.schedule.shifts)
        total_people = len(self.schedule.people)
        
        for person_name, person in self.schedule.people.items():
            assigned_hours = sum(self.A[(person_name, shift_idx)] * shift.hours for shift_idx, shift in enumerate(self.schedule.shifts))
            error[person_name] = self.model.addVar(lb=-np.inf, ub=np.inf, name=f"error_{person_name}")

            person.expected_share = total_shift_hours/total_people  # For now even distribution between shifts

            self.model.addConstr(error[person_name] == person.expected_share - assigned_hours, name=f"C1_Error_{person_name}")
            objective += error[person_name] ** 2
        
        # Minimize objective while considering unfilled shifts (via slack variable)
        objective += sum(self.slack[shift_idx] for shift_idx in self.slack) * 1000 # Large penalty for no assignment 
        
        self.model.setObjective(objective, gp.GRB.MINIMIZE)

    def apply_constraints(self):
        self._apply_shift_assignment_constraints()
        self._apply_availability_constraints()
        self._apply_no_night_to_morning_constraints()
        self._apply_no_evening_to_morning_constraints()
        self._apply_no_sunday_if_quota_not_met_constraints()
        self._apply_max_one_sunday_shift_constraints()
        self._apply_max_hours_constraints()

    def _apply_shift_assignment_constraints(self):
        for shift_idx, shift in enumerate(self.schedule.shifts):
            # Total assignments for each shift must equal the number of persons required for that shift, plus slack for unfilled shifts
            self.model.addConstr(
                sum(self.A[(person_name, shift_idx)] for person_name in self.schedule.people) + self.slack[shift_idx],
                gp.GRB.EQUAL, 
                shift.persons_required, 
                name=f"C2_ShiftAssignment_{shift_idx}"
            )

    def _apply_availability_constraints(self):
        for shift_idx, shift in enumerate(self.schedule.shifts):
            for person_name, person in self.schedule.people.items():
                self.model.addConstr(self.A[(person_name, shift_idx)] <= person.availability[shift_idx],
                                     name=f"C3_Availability_{person_name}_{shift_idx}")

    def _apply_no_night_to_morning_constraints(self):
        for person_name, person in self.schedule.people.items():
            prev_shift = None
            prev_type = None
            for shift_idx, shift in enumerate(self.schedule.shifts):
                if prev_shift is not None and prev_type == "Avond" and shift.shift_type == "Ochtend":
                    self.model.addConstr(self.A[(person_name, shift_idx)] + self.A[(person_name, prev_shift)] <= 1,
                                         name=f"C4_NoNightToMorning_{person_name}_{shift_idx}")
                prev_shift, prev_type = shift_idx, shift.shift_type

    def _apply_no_evening_to_morning_constraints(self):
        for person_name, person in self.schedule.people.items():
            prev_shift, prev_prev_shift = None, None
            prev_type, prev_prev_type = None, None
            for shift_idx, shift in enumerate(self.schedule.shifts):
                if shift.shift_type == "Avond":
                    if prev_type in ["Middag", "Ochtend"]:
                        self.model.addConstr(self.A[(person_name, prev_shift)] + self.A[(person_name, shift_idx)] <= 1,
                                             name=f"C5.1_NoAfternoonToEvening_{person_name}_{shift_idx}")
                    if prev_prev_type == "Ochtend":
                        self.model.addConstr(self.A[(person_name, prev_prev_shift)] + self.A[(person_name, shift_idx)] <= 1,
                                             name=f"C5.2_NoMorningToEvening_{person_name}_{shift_idx}")
                prev_prev_shift, prev_prev_type = prev_shift, prev_type
                prev_shift, prev_type = shift_idx, shift.shift_type

    def _apply_no_sunday_if_quota_not_met_constraints(self):
        for person_name, person in self.schedule.people.items():
            if person.non_sunday_hours < self.sunday_quota:
                for shift_idx, shift in enumerate(self.schedule.shifts):
                    if shift.day == "Zondag":
                        self.model.addConstr(self.A[(person_name, shift_idx)] == 0,
                                             name=f"C6_NoSundayIfQuotaNotMet_{person_name}_{shift_idx}")

    def _apply_max_one_sunday_shift_constraints(self):
        for person_name in self.schedule.people:
            self.model.addConstr(sum(self.A[(person_name, shift_idx)] for shift_idx, shift in enumerate(self.schedule.shifts) if shift.day == "Zondag") <= 1,
                                 name=f"C7_OneSundayShift_{person_name}")

    def _apply_max_hours_constraints(self):
        for person_name in self.schedule.people:
            self.model.addConstr(sum(self.A[(person_name, shift_idx)] * shift.hours for shift_idx, shift in enumerate(self.schedule.shifts)) <= self.max_hours,
                                 name=f"C8_MaxHours_{person_name}")

    def solve(self):
        self.model.optimize()
