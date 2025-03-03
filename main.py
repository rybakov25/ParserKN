# -*- coding: utf-8 -*-
# Python version 3.12
# --------------------------------------------------------------┐
# Name:        ParserKN-pdf                                     |
#                                                               |
# Author:      Pavel Rybakov                                    |
#                                                               |
# Copyright:   © Pavel Rybakov 2025                             |
# Licence:     MIT                                              |
# --------------------------------------------------------------┘
__author__ = "Pavel Rybakov"
__copyright__ = "2025 Pavel Rybakov"
__license__ = "MIT"

# Import libraries
import os
import re
import sys
import json
import csv
import pdf2docx
import pandas as pd
from typing import NoReturn
from datetime import timedelta
from transliterate.base import registry
from transliterate.discover import autodiscover
from transliterate import translit, get_available_language_codes
from transliterate.exceptions import LanguageCodeError, InvalidRegistryItemType
from tkinter import filedialog

# Import my libraries
from filter_list import unnecessary_values
from programs_data import ProgramsData, ProgramsDataEncoder
from MyTranslit import MyTranslit

TEMP_FILE = "./tmp/last_output.docx"
try:
    registry.register(MyTranslit)
    autodiscover()
    if MyTranslit.language_code in get_available_language_codes():
        pass
except LanguageCodeError as lce:
    print(lce)
    sys.exit(1)
except InvalidRegistryItemType as irit:
    print(irit)
    sys.exit(1)


class Parser:
    """Parser pdf files"""

    def __init__(self, pdf_file_path: str, output_type: str = "json") -> NoReturn:
        self.__pdf_file: str = pdf_file_path.lower()
        self.__output_type: str = output_type
        self.__output_file: str = self.__pdf_file.split("/")[-1].split(".")[0] + f".{output_type}"
        self.__tables: list = []
        self.__data: ProgramsData = ProgramsData()
        self.__cur_prog: str = ""
        self.__cur_tool: str = ""
        self.__time_cur_tool: str = ""
        self.__get_pdf_tables()

    def parse(self) -> NoReturn:
        """Parsing extracted tables"""
        self.__data.date = f"{self.__tables[0][2][1].strip()} {self.__tables[0][3][1].strip()}"
        self.__row_parse(self.__tables[0][-1])
        for table in self.__tables[3:]:
            for row in table:
                self.__row_parse(row)
        match self.__output_type:
            case "json":
                self.to_json()
            case "csv":
                self.get_csv()
            case "xlsx":
                self.get_xlsx()
            case _:
                raise AttributeError("Field __output_type is non valid")

    def __row_parse(self, row: list) -> NoReturn:
        mid_row = list(filter(lambda x: x is not None and bool(x), row))
        new_row = list(map(lambda x: x.strip(), mid_row))
        if self.__filter(new_row):
            return
        if "Общее" in new_row[0]:
            if not self.__data.total_time:
                self.__data.total_time = new_row[1]
            else:
                self.__total_time_update(new_row[1])
            return
        for cell in new_row:
            cipher = re.search(r"\d?[К-Ш]+-\d{1,4}", cell, re.U)
            program = re.search(r"(\.MPF|\.NC)", cell, re.I | re.U)
            cur_tool = re.search(r"[А-Я]?[а-я]+ [D-Z]+", cell, re.U)
            tool_time = re.search(r"\d+:\d+:\d+", cell, re.U)
            if cipher and cipher[0] in cell:
                self.__data.cipher = cell
            if program and program[0] in cell:
                self.__cur_prog = cell
                self.__data.programs.update({cell: dict()})
            if cur_tool and cur_tool[0] in cell:
                self.__cur_tool = cell
                if self.__cur_tool in self.__data.programs[self.__cur_prog]:
                    continue
                self.__data.programs[self.__cur_prog].update({cell: ""})
            if tool_time and tool_time[0] in cell:
                if self.__data.programs[self.__cur_prog][self.__cur_tool]:
                    self.__time_update(cell)
                else:
                    self.__data.programs[self.__cur_prog][self.__cur_tool] = cell

    def __time_update(self, cell: str) -> NoReturn:
        oh, o_min, o_sec = list(map(int, self.__data.programs[self.__cur_prog][self.__cur_tool].split(":")))
        nh, n_min, n_sec = list(map(int, cell.split(":")))
        sum_time = timedelta(hours=oh, minutes=o_min, seconds=o_sec)
        sum_time += timedelta(hours=nh, minutes=n_min, seconds=n_sec)
        self.__data.programs[self.__cur_prog][self.__cur_tool] = str(sum_time)

    def __total_time_update(self, cell: str) -> NoReturn:
        oh, o_min, o_sec = list(map(int, self.__data.total_time.split(":")))
        nh, n_min, n_sec = list(map(int, cell.split(":")))
        sum_time = timedelta(hours=oh, minutes=o_min, seconds=o_sec)
        sum_time += timedelta(hours=nh, minutes=n_min, seconds=n_sec)
        ts = sum_time.total_seconds()
        h = int(ts // 3600)
        m = int((ts % 3600) // 60)
        s = int(ts % 60)
        self.__data.total_time = f"{h}:{m}:{s}"

    def __filter(self, row: list) -> bool:
        """Filtering rows"""
        if not all(row):
            return True
        s = " ".join(row).lower()
        volume = re.search(r"\d*,\d", s) if len(row) == 2 else 0
        for value in unnecessary_values:
            if (value in s) or (volume and volume[0] in s):
                return True
        return False

    def __get_pdf_tables(self) -> NoReturn:
        """Gets tables from pdf file"""
        doc = pdf2docx.Converter(self.__pdf_file)
        try:
            parse_settings = doc.default_settings
            cpu_count = os.cpu_count() // 2  # Don't get involved if you don't know what you're doing.
            # -------------------------------------------------------------------------------------------------
            # The number of processor cores used for parsing.
            # >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> Warning: <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            # If the specified number is greater than the number of physically available cores,
            # there is a risk that the computer will freeze "dead".
            parse_settings["cpu_count"] = cpu_count
            # -------------------------------------------------------------------------------------------------
            # parse_settings["multi_processing"] = True
            self.__tables: list = doc.extract_tables(0, None, None, **parse_settings)
            if len(sys.argv) > 1 and sys.argv[1] == "wod":
                doc.convert(TEMP_FILE)
        except pdf2docx.converter.ConversionException as ce:
            print(ce)
        finally:
            doc.close()

    def to_json(self) -> NoReturn:
        """Preparing to convert to json"""
        res_data: dict = {}
        res_data.update({"cipher": translit(self.__data.cipher, "ru_nal")})
        for prog in self.__data.programs:
            res_data.update({prog: dict()})
            for tool, time in self.__data.programs[prog].items():
                tool = translit(tool, "ru_nal")
                res_data[prog].update({tool: time})
        res_data.update({"total_time": self.__data.total_time})
        self.get_json(res_data)

    def get_json(self, data: dict) -> NoReturn:
        """Converting extracting datas from pdf file to json file"""
        file_path: str = f"{os.getenv('USERPROFILE')}\\Desktop\\{self.__output_file}"
        with open(file_path, "w") as json_file:
            json.dump(data, json_file, cls=ProgramsDataEncoder, indent=4)

    def get_csv(self) -> NoReturn:
        file_path: str = f"{os.getenv('USERPROFILE')}\\Desktop\\{self.__output_file}"
        data_dict: dict = self.__data.to_dict()
        with open(file_path, "w", newline="") as csv_file:
            writer = csv.writer(csv_file, delimiter=";")
            writer.writerow(["Дата отчета", data_dict["date"]])
            writer.writerow(["Шифр", data_dict["cipher"]])
            for _, (prog_id, tools) in enumerate(data_dict["programs"].items()):
                writer.writerow([prog_id])
                for _, (tool, time) in enumerate(tools.items()):
                    writer.writerow([tool, time])
            writer.writerow(["Общее время", data_dict["total_time"]])

    def get_xlsx(self):
        # todo: finish write, don't use in work
        df = pd.DataFrame().from_dict(self.__data.to_dict(), orient="columns")
        self.__output_file = f"{os.getenv('USERPROFILE')}\\Desktop\\{self.__output_file}"
        df.to_excel(self.__output_file, index=False, sheet_name=self.__data.cipher)


def main() -> NoReturn:
    try:
        pdf_paths = filedialog.askopenfilenames(initialdir=f"{os.getenv('USERPROFILE')}/Desktop", title="Choice pdf's")
        ot: str = "csv"
        if len(pdf_paths) == 1:
            parser = Parser(pdf_paths[0], ot)
            parser.parse()
        else:
            for pdf_path in pdf_paths:
                parser = Parser(pdf_path, ot)
                parser.parse()
    except ValueError as ve:
        print(ve)
        sys.exit(1)


if __name__ == "__main__":
    main()
