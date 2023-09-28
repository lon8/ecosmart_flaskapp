import pandas as pd


class Database:
    CONFIG_SHEET = "Config"
    PROJECTS_SHEET = "Projects"
    PROJECT_PATH_COL = "ProjectPath"
    DEFAULT_PROJECTS_PATH = "DefaultProjectsFolder"
    PROJECT_HEADER_ROW = 3

    def __init__(self, path):
        self._path = path

    def get_config(self):
        result = {}
        df = pd.read_excel(self._path, sheet_name=Database.CONFIG_SHEET, header=None, index_col=0)

        for k, v in df.to_dict("index").items():
            result[k] = v[1]

        return result

    def get_projects(self):
        df = pd.read_excel(
            self._path, sheet_name=Database.PROJECTS_SHEET, header=Database.PROJECT_HEADER_ROW, keep_default_na=False
        )

        return df.to_dict("records"), list(df.columns)

    def write_projects(self, values):
        df = pd.DataFrame.from_dict(values)

        with pd.ExcelWriter(self._path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            df.to_excel(writer, index=False, sheet_name=Database.PROJECTS_SHEET, startrow=Database.PROJECT_HEADER_ROW)

    def find_project(self, projects, name):
        project_name_lc = name.lower()
        return next(filter(lambda p: p["ProjectName"].lower() == project_name_lc, projects), None)