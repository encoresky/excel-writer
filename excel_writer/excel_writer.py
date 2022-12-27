import os
import pandas as pd
from typing import Dict, List, Union

class ExcelWriter():
    def __init__(self, file_name: str = None,
                        sheet_name:str = "Sheet1") -> None:
        """Initiates the Excel Writer

        Args:
            file_name (str, optional): file name/path in which you want to perform 
                                        operations with extension. Defaults to None.
            sheet_name (str, optional): name of the sheet in the file
        """
        self.file_name = file_name
        self.sheet_name = sheet_name
        self.style_properties = {
            'text-align': 'center'
        }

    def concat_new_data(self, data_dict: Dict = {},
                              data_df: pd.DataFrame = None) -> pd.DataFrame:
        """concatenates the new data in the old data

        Args:
            data_dict (Dict): data to append in the file
            data_df (pd.DataFrame): data to append in the file

        Returns:
            pd.DataFrame: Returns the data of the file
        """
        new_df = pd.DataFrame(data_dict)
        if issubclass(pd.DataFrame, type(data_df)):
            new_df = data_df

        old_df = self.read_data_in_pd()

        df = pd.concat([old_df, new_df], ignore_index=True)
        return df

    def write_data(self, data_dict: Dict = {},
                         data_df: pd.DataFrame = None,
                         update_data: bool = True) -> None:
        """writes the input data

        Args:
            data_dict (Dict): data to write in the file
            data_df (pd.DataFrame): data to write in the file
        """
        df = pd.DataFrame(data_dict)
        if issubclass(pd.DataFrame, type(data_df)):
            df = data_df

        if update_data:
            df = self.concat_new_data(data_dict)
            if issubclass(pd.DataFrame, type(data_df)):
                df = self.concat_new_data(data_df=data_df)

        with pd.ExcelWriter(self.file_name, engine='openpyxl') as writer:
            df.style.set_properties(**self.style_properties).to_excel(writer, sheet_name=self.sheet_name, index=False)

    def read_data_in_dict(self) -> Dict:
        """reads the data from file

        Returns:
            Dict: Returns the data of the file
        """
        if not os.path.exists(self.file_name):
            return pd.DataFrame().to_dict(orient='list')

        df = pd.read_excel(self.file_name)
        data = df.to_dict(orient='list')

        return data

    def read_data_in_pd(self) -> pd.DataFrame:
        """reads the data from file

        Returns:
            pd.DataFrame: Returns the data of the file
        """
        if not os.path.exists(self.file_name):
            return pd.DataFrame()

        df = pd.read_excel(self.file_name)
        return df

    def add_column(self, col_name: str, 
                         position: int = None,
                         values: List[str] = None) -> None:
        """Adds a column in an existing file

        Args:
            col_name (str): Name of the column
            position (int, Optional): Position/Index of the column
            values (List[str], Optional): input values for rows
        """
        old_df = self.read_data_in_pd()
        values = values if values else ""

        if position:
            old_df.insert(position, col_name, values)
        else:
            old_df[col_name] = values

        self.write_data(data_df=old_df, update_data=False)
        
    def get_column(self, col_name: Union[List[str], str],
                         return_type: str = "dict") -> Union[Dict, pd.DataFrame]:
        """get the selected column details

        Args:
            col_name (Union[List[str], str]): column name for which you want details
            return_type (str, optional): format of data you want to return. Defaults to "dict"
                                         Options ("dict", "df")

        Returns:
            Union[Dict, pd.DataFrame]: Returns the data of selected column
        """

        if type(col_name) == str:
            col_name = [col_name]

        df = self.read_data_in_pd()
        column = df[col_name]

        if return_type == "df":
            return column

        return column.to_dict(orient='list')

    def delete_column(self, name: Union[List[str], str],
                            return_type: str = "dict",
                            write_in_file: bool = False) -> None:
        """deletes the column/columns from the data

        Args:
            name (Union[List[str], str]): name of the column to delete
            return_type (str, optional): format of data you want to return. Defaults to "dict"
                                         Options ("dict", "df")
            write_in_file (bool) -> if you want to write in current working file (default False)
        """

        if type(name) == str:
            name = [name]

        df = self.read_data_in_pd()
        df.drop(name, axis=1, inplace=True)

        if write_in_file:
            self.write_data(data_df=df, update_data=False)

        if return_type == "df":
            return df

        return df.to_dict(orient='list')

    def add_row(self, position: int = None,
                      values: Dict = None) -> None:
        """Adds a column in an existing file

        Args:
            position (int, Optional): Position/Index of the row
            values (Dict, Optional): input values for rows {col_name: value}
        """
        old_df = self.read_data_in_pd()

        data = dict.fromkeys(old_df.columns, [""])
        if values:
            for k, _ in data.items():
                if k in values:
                    data.update({
                        k: [values[k]]
                    })
                else:
                    data.update({
                        k: [""]
                    })

        if position:
            old_df.loc[position] = [each for values in data.values() 
                                                        for each in values]
            df = old_df
        else:
            new_df = pd.DataFrame(data)
            df = new_df

        self.write_data(data_df=df, update_data=False)

    def get_row(self, position: Union[List[int], int],
                      return_type: str = "dict") -> Union[Dict, pd.DataFrame]:
        """get the selected column details

        Args:
            position (Union[List[int], int]): Position/Index of the row
            return_type (str, optional): format of data you want to return. Defaults to "dict"
                                         Options ("dict", "df")

        Returns:
            Union[Dict, pd.DataFrame]: Returns the data of selected column
        """

        if type(position) == int:
            position = [position]

        df = self.read_data_in_pd()
        column = df.loc[position]

        if return_type == "df":
            return column

        return column.to_dict(orient='list')

    def delete_row(self, position: Union[List[int], int],
                         return_type: str = "dict",
                         write_in_file: bool = False) -> None:
        """deletes the row/rows from the data

        Args:
            position (Union[List[int], int]): position of the row to delete
            return_type (str, optional): format of data you want to return. Defaults to "dict"
                                         Options ("dict", "df")
            write_in_file (bool) -> if you want to write in current working file (default False)
        """

        if type(position) == int:
            position = [position]

        df = self.read_data_in_pd()
        df.drop(position, axis=0, inplace=True)

        if write_in_file:
            self.write_data(data_df=df, update_data=False)

        if return_type == "df":
            return df

        return df.to_dict(orient='list')

    def get_cell(self, row_index: int,
                       col_name: str) -> str:
        """get the value at specific cell 

        Args:
            row_index (int): Position/Index of the row
            col_name (str): name of the column

        Returns:
            str: the value at specific cell
        """

        df = self.read_data_in_pd()
        cell = df.at[row_index, col_name]
        
        return cell

    def update_cell(self, row_index: int,
                          col_name: str,
                          value: str,
                          return_type: str = "dict",
                          write_in_file: bool = False) -> None:
        """update the value at specific cell 

        Args:
            row_index (int): Position/Index of the row
            col_name (str): name of the column
            value (str): value to insert into the cell
            return_type (str, optional): format of data you want to return. Defaults to "dict"
                                         Options ("dict", "df")
            write_in_file (bool) -> if you want to write in current working file (default False)
        """

        df = self.read_data_in_pd()
        df.at[row_index, col_name] = value
        
        if write_in_file:
            self.write_data(data_df=df, update_data=False)

        if return_type == "df":
            return df

        return df.to_dict(orient='list')

    def delete_cell(self, row_index: int,
                          col_name: str,
                          return_type: str = "dict",
                          write_in_file: bool = False) -> None:
        """deletes the value at specific cell 

        Args:
            row_index (int): Position/Index of the row
            col_name (str): name of the column
            return_type (str, optional): format of data you want to return. Defaults to "dict"
                                         Options ("dict", "df")
            write_in_file (bool) -> if you want to write in current working file (default False)
        """

        df = self.read_data_in_pd()
        df.at[row_index, col_name] = ""
        
        if write_in_file:
            self.write_data(data_df=df, update_data=False)

        if return_type == "df":
            return df

        return df.to_dict(orient='list')

    def add_sheet(self, sheet_name: str,
                        data_dict: Dict = {},
                        data_df: pd.DataFrame = None) -> None:
        """adds the sheet in the file

        Args:
            sheet_name (str): name of the sheet in the file
            data_dict (Dict): data to write in the file
            data_df (pd.DataFrame): data to write in the file
        """
        
        df = pd.DataFrame(data_dict)
        if issubclass(pd.DataFrame, type(data_df)):
            df = data_df

        with pd.ExcelWriter(self.file_name, engine='openpyxl', mode="a") as writer:
            df.style.set_properties(**self.style_properties).to_excel(writer, sheet_name=sheet_name, index=False)

    def update_sheet(self, sheet_name: str,
                           data_dict: Dict = {},
                           data_df: pd.DataFrame = None) -> None:
        """updates the sheet in the file

        Args:
            sheet_name (str): name of the sheet in the file
            data_dict (Dict): data to write in the file
            data_df (pd.DataFrame): data to write in the file
        """
        
        new_sheet = pd.DataFrame(data_dict)
        if issubclass(pd.DataFrame, type(data_df)):
            new_sheet = data_df

        file = pd.ExcelFile(self.file_name)
        old_sheet = pd.read_excel(file, sheet_name)

        df = pd.concat([old_sheet, new_sheet], ignore_index=True)
        self.delete_sheet(sheet_name)

        with pd.ExcelWriter(self.file_name, engine='openpyxl', mode="a") as writer:
            df.style.set_properties(**self.style_properties).to_excel(writer, sheet_name=sheet_name, index=False)

    def delete_sheet(self, sheet_name: str) -> None:
        """removes the sheet from the file

        Args:
            sheet_name (str): name of the sheet in the file
        """
        with pd.ExcelWriter(self.file_name, engine='openpyxl', mode='a') as writer: 
            workBook = writer.book
            try:
                workBook.remove(workBook[sheet_name])
            except:
                print("There is no such sheet in this file")