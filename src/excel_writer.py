import pandas as pd
import xlwt
from util import Schedule, Shift, Person

class ExcelTool:
    @staticmethod
    def read_availability(file_path):
        availability = pd.read_excel(file_path, header=1, sheet_name="Hele Team")
        availability = availability.iloc[:-5]  # Drop last 5 rows containing no useful info
        schedule = Schedule()
        people = availability.iloc[0:1, 13:27].columns

        for person_name in people:
            person = Person(person_name)
            schedule.people[person_name] = person

        for _, row in availability.iterrows():
            shift = Shift(
                row["Poule Library"],
                ExcelTool.calc_hours(row["Poule Library"]),
                row["Benodigd (lib)"],
                row["Type"],
                row["Dag"],
                row["Datum"],
            )

            schedule.shifts.append(shift)

            for person_name, person in schedule.people.items():
                available = row[person_name] in ["j", "J", "x", "X"]
                person.availability.append(1 if available else 0)

        return schedule

    @staticmethod
    def write_schedule(schedule, solver, filename):
        """Writes the optimized schedule to an Excel file with shifts and a summary."""
        data = []
        # Reset previous assignments and bonuses
        for person in schedule.people.values():
            person.assigned_shifts = []
            person.bonus_hours = 0.0

        for shift_idx, shift in enumerate(schedule.shifts):
            assigned_people = [person_name for person_name in schedule.people.keys() if solver.A[(person_name, shift_idx)].X > 0.5]

            # Use the bonus_hours attribute of the Shift class
            bonus = shift.bonus_hours

            # Append shift data with bonus hours
            data.append([shift.day, shift.date, shift.shift_type, shift.hours, bonus] + assigned_people)

            # Update assigned shifts and calculate bonus hours for each person
            for person_name in assigned_people:
                person = schedule.people[person_name]
                person.assigned_shifts.append(shift)
                person.bonus_hours += bonus

        # Prepare shifts DataFrame
        max_people = max((len(row) - 5 for row in data), default=0)
        columns = ["Day", "Date", "Shift Type", "Hours", "Bonus Hours"] + [f"Person {i+1}" for i in range(max_people)]
        df_shifts = pd.DataFrame(data, columns=columns)

        # Prepare summary DataFrame
        summary_data = []
        for person_name, person in schedule.people.items():
            total_hours = sum(shift.hours for shift in person.assigned_shifts)
            summary_data.append([person_name, total_hours, person.bonus_hours])
        df_summary = pd.DataFrame(summary_data, columns=["Name", "Total Hours", "Bonus Hours"])

        # Write to Excel with multiple sheets
        with pd.ExcelWriter(filename) as writer:
            df_shifts.to_excel(writer, sheet_name="Shifts", index=False)
            df_summary.to_excel(writer, sheet_name="Summary", index=False)

    @staticmethod
    def write_metrics(schedule, solver, filename):
        """Writes the metrics calculated from the solver to an Excel file."""
        metrics = {}
        days_order = ["Maandag", "Dinsdag", "Woensdag", "Donderdag",
                    "Vrijdag", "Zaterdag", "Zondag"]
        spread_penalty_coeff = 0.1
        distance_penalties = {1: 0.3, 2: 0.2, 3: 0.1}
        
        # Calculate distribution penalty
        total_distribution_penalty = 0
        
        # Existing metrics
        total_slack = sum(slack.X for slack in solver.slack.values())
        total_required_persons = sum(shift.persons_required for shift in schedule.shifts)
        filled_positions = total_required_persons - total_slack

        if total_required_persons > 0:
            metrics['Filled Shifts (%)'] = (filled_positions / total_required_persons) * 100
        else:
            metrics['Filled Shifts (%)'] = 0

        total_regular_error = 0
        total_bonus_error = 0
        
        # New: Pre-calculate day indexes for all shifts
        day_indexes = {}
        for shift in schedule.shifts:
            day = str(shift.day).strip().title()
            day_indexes[shift] = days_order.index(day)

        for name, person in schedule.people.items():
            # Existing calculations
            regular_hours = sum(solver.A[(name, i)].X * shift.hours for i, shift in enumerate(schedule.shifts))
            bonus_hours = sum(solver.A[(name, i)].X * shift.bonus_hours for i, shift in enumerate(schedule.shifts))
            
            expected_regular = (person.available_regular_hours / schedule.get_total_available_regular()) * sum(shift.hours * shift.persons_required for shift in schedule.shifts)
            expected_bonus = (person.available_regular_hours / schedule.get_total_available_regular()) * sum(shift.bonus_hours * shift.persons_required for shift in schedule.shifts)
            
            total_regular_error += (expected_regular - regular_hours) ** 2
            total_bonus_error += (expected_bonus - bonus_hours) ** 2

            # Distribution penalty calculation
            person_penalty = 0
            assigned_shifts = person.assigned_shifts
            
            # Compare all pairs of shifts
            for i in range(len(assigned_shifts)):
                for j in range(i+1, len(assigned_shifts)):
                    shift1 = assigned_shifts[i]
                    shift2 = assigned_shifts[j]
                    
                    # Get pre-calculated day indexes
                    idx1 = day_indexes[shift1]
                    idx2 = day_indexes[shift2]
                    
                    # Calculate day gap with weekly wrap-around
                    day_diff = min(abs(idx2 - idx1), 7 - abs(idx2 - idx1))
                    
                    if day_diff in distance_penalties:
                        person_penalty += distance_penalties[day_diff]
            
            # Apply global coefficient
            person_penalty *= spread_penalty_coeff
            total_distribution_penalty += person_penalty
            metrics[f'{name} Distribution Penalty'] = person_penalty

            # Existing share calculations
            received_regular_share = regular_hours / person.available_regular_hours if person.available_regular_hours > 0 else 0
            expected_regular_share = expected_regular / person.available_regular_hours if person.available_regular_hours > 0 else 0
            received_bonus_share = bonus_hours / person.available_bonus_hours if person.available_bonus_hours > 0 else 0
            expected_bonus_share = expected_bonus / person.available_bonus_hours if person.available_bonus_hours > 0 else 0
            received_total_share = (regular_hours + bonus_hours) / person.total_available_hours if person.total_available_hours > 0 else 0
            expected_total_share = (expected_regular + expected_bonus) / person.total_available_hours if person.total_available_hours > 0 else 0

            # Individual metrics
            metrics[f'{name} Regular Hours'] = regular_hours
            metrics[f'{name} Bonus Hours'] = bonus_hours
            metrics[f'{name} Received Regular Share'] = received_regular_share
            metrics[f'{name} Expected Regular Share'] = expected_regular_share
            metrics[f'{name} Received Bonus Share'] = received_bonus_share
            metrics[f'{name} Expected Bonus Share'] = expected_bonus_share
            metrics[f'{name} Received Total Share'] = received_total_share
            metrics[f'{name} Expected Total Share'] = expected_total_share

        # Add fairness metrics
        metrics['Regular Hours Fairness (SSE)'] = total_regular_error
        metrics['Bonus Hours Fairness (SSE)'] = total_bonus_error
        metrics['Total Distribution Penalty'] = total_distribution_penalty

        # Convert metrics to DataFrame
        df_metrics = pd.DataFrame(list(metrics.items()), columns=['Metric', 'Value'])

        # Write metrics to Excel
        with pd.ExcelWriter(filename) as writer:
            df_metrics.to_excel(writer, sheet_name="Metrics", index=False)

    @staticmethod
    def calc_hours(time_str):
        h_start, m_start, h_end, m_end = (
            int(time_str[:2]),
            int(time_str[3:5]),
            int(time_str[6:8]),
            int(time_str[9:11]),
        )
        return ((h_end * 60 + m_end) - (h_start * 60 + m_start)) / 60