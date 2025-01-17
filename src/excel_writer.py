import pandas as pd
import xlwt
from util import Schedule
from util import Shift
from util import Person

class ExcelTool:
    @staticmethod
    def read_availability(file_path):
        availability = pd.read_excel(file_path, header=1, sheet_name="Hele Team")
        availability = availability.iloc[:-5] # Drop last 5 rows containing no useful info 
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

        # print(schedule)

        return schedule

    @staticmethod
    def write_schedule(schedule, solver, filename):
        """Writes the optimized schedule to an Excel file in the same format as the original roster."""
        data = []
        for shift_idx, shift in enumerate(schedule.shifts):
            assigned_people = [person_name for person_name in schedule.people.keys() if solver.A[(person_name, shift_idx)].X > 0.5]
            data.append([shift.day, shift.date, shift.shift_type, shift.hours] + assigned_people)

        # Convert to DataFrame
        max_people = max(len(row) - 4 for row in data)  # Find max assigned people per shift
        columns = ["Day", "Date", "Shift Type", "Hours"] + [f"Person {i+1}" for i in range(max_people)]
        df = pd.DataFrame(data, columns=columns)
        
        # Save to Excel
        df.to_excel(filename, index=False)


    @staticmethod
    def calc_hours(time_str):
        h_start, m_start, h_end, m_end = (
            int(time_str[:2]),
            int(time_str[3:5]),
            int(time_str[6:8]),
            int(time_str[9:11]),
        )
        return ((h_end * 60 + m_end) - (h_start * 60 + m_start)) / 60
