using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;


namespace TricentisExtensions.Modules.CustomEngine
{
    public class AOOperations
    {
        private readonly string _dataSource;

        private readonly Excel.Application xlApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

        private enum Functions
        {
            GetVariables,
            GetFilters
        }

        public AOOperations(string dataSource = "DS_1")
        {
            _dataSource = dataSource;
        }

        public Excel.Application GetXlApp()
        {
            return xlApp; 
        }

        private Dictionary<string, string> GetFunctionResult(Functions function, string param = "")
        {
            Array _tmpResultArray;
            var dict = new Dictionary<string, string>();

            switch (function)
            {
                case Functions.GetVariables:
                    _tmpResultArray = (Array)(object)xlApp.Run("SAPListOfVariables", _dataSource, "Key", param.ToUpper());
                    break;
                case Functions.GetFilters:
                    var funcResult = xlApp.Run("SAPListOfDynamicFilters", _dataSource, "Key");
                    if (String.IsNullOrEmpty(funcResult))
                        _tmpResultArray = Array.Empty<string>();
                    else
                        _tmpResultArray = (Array)(object)funcResult;
                    break;
                default:
                    _tmpResultArray = Array.Empty<string>();
                    break;
            }
            try
            {
                if (_tmpResultArray.Rank > 1)
                {
                    string[] key = new string[_tmpResultArray.GetLength(0)];
                    string[] value = new string[_tmpResultArray.GetLength(0)];
                    var _tempFiltersArray = (object[,])_tmpResultArray;

                    for (int i = 1; i <= _tmpResultArray.GetLength(0); i++)
                    {
                        for (int j = 1; j <= _tmpResultArray.GetLength(1); j++)
                        {
                            switch (j)
                            {
                                case 1:
                                    key[i - 1] = (string)_tempFiltersArray[i, j];
                                    break;
                                case 2:
                                    value[i - 1] = (string)_tempFiltersArray[i, j];
                                    break;
                                default:
                                    break;
                            }
                        }
                        dict.Add(key[i - 1], value[i - 1]);
                    }
                }
                else if (_tmpResultArray.Length == 0)
                {
                    dict = new Dictionary<string, string> 
                    {
                        { string.Empty, string.Empty}
                    };
                }
                else
                {
                    dict.Add((string)_tmpResultArray.GetValue(1), (string)_tmpResultArray.GetValue(2));
                }

                return dict;
            }
            catch (InvalidCastException e)
            {
                throw e;
            }

        }


        public int LogIn(string system, string userName, string password)
        {
            return (int)xlApp.Run("SAPLogOn", _dataSource, system, userName, password);
        }

        public int LogOff()
        {
            return (int)xlApp.Run("SAPLogOff", false);
        }

        public string GetVariableTechnicalName(string varName)
        {
            return xlApp.Run("SAPGetVariable", _dataSource, varName, "TECHNICALNAME");
        }

        public int SetVariable(string prompt, string value)
        {
            return (int)xlApp.Run("SAPSetVariable", GetVariableTechnicalName(prompt), value, "INPUT_STRING", _dataSource);
        }       

        public int SetVariables(Dictionary<string, string> dict)
        {
            int i = 0;

            foreach (var pair in dict)
            {
                i = SetVariable(pair.Key, pair.Value);
            }

            return i;
        }

        public Dictionary<string, string> GetVariables(string param = "PROMPTS")
        {
            /*
             * Params:
             *  ALL to display all variables (filled and unfilled) including variables not visible on the prompts dialog.
                PROMPTS to display all variables (filled and unfilled) visible on the prompts dialog.
                ALL_FILLED to display all filled variables including variables not visible on the prompts dialog.
                PROMPTS_FILLED to display all filled variables visible on the prompts dialog.
                PLAN_PARAMETER to display all variables (filled and unfilled) of a planning object.
             */
            return GetFunctionResult(Functions.GetVariables, param.ToUpper());
            
        }
        

        public Dictionary<string, string> GetActiveFilters()
        {
            return GetFunctionResult(Functions.GetFilters);
        }

        public List<string> GetActiveFilterNames()
        {
            return GetActiveFilters().Keys.Where(key => !key.ToUpper().Equals("MEASURES")).ToList();
        }

        public List<string> GetActiveFilterValues()
        {
            return GetActiveFilters().Values.Where(key => !key.ToUpper().Equals("MEASURES")).ToList();
        }

        public List<string> GetMeasures()
        {
            //var measures = GetActiveFilters().Where(k => k.Key.ToUpper().Equals("MEASURES")).Select(val => val.Value);
            var measures = xlApp.Run("SAPGetDisplayedMeasures", _dataSource);
            var values = new List<string>();

            //foreach (var item in measures)
            //{
            //    if (!string.IsNullOrEmpty(item))
            //    {
            //        values.AddRange(item.Split(';'));
            //    }
            //}

            if (!string.IsNullOrEmpty(measures))
            {
                values.AddRange(measures.Split(';'));
            }

            values = values.Select(x => x.Trim()).ToList();

            return values;
        }

        public int ResetAllFilters()
        {
            var filters = GetActiveFilterNames();

            PauseRefresh();
            foreach (var filter in filters)
            {
                SetFilter(filter, "ALLMEMBERS");
            }
            ActivateRefresh();
            filters = GetActiveFilterNames();

            if (filters.Count == 0)
            {
                return 1;
            }
            return 0;
        }


        public string[] GetDimensions()
        {
            var dimensionArray = xlApp.Run("SAPListOfDimensions", _dataSource, "Description");

            string[] value = new string[dimensionArray.GetLength(0)];


            for (int i = 1; i <= dimensionArray.GetLength(0); i++)
            {
                value[i - 1] = (string)dimensionArray[i, 2];
            }

            return value;
        }

        public string GetDimensionTechName(string dimName)
        {
            object[,] dimensionArray = xlApp.Run("SAPListOfDimensions", _dataSource, "Description");


            for (int i = 1; i <= dimensionArray.GetLength(0); i++)
            {
                if ((((string)dimensionArray[i, 2]).ToUpper()).Equals(dimName.ToUpper()))
                {
                    return (string)dimensionArray[i, 1];
                }
            }
            return "";
        }

        public string[] GetDimensionTechNames()
        {
            var dimensionArray = xlApp.Run("SAPListOfDimensions", _dataSource, "Description");
            string[] value = new string[dimensionArray.GetLength(0)];


            for (int i = 1; i <= dimensionArray.GetLength(0); i++)
            {
                value[i - 1] = (string)dimensionArray[i, 1];
            }
            return value;
        }

        public List<string> GetColumns()
        {
            var dimensionArray = xlApp.Run("SAPListOfDimensions", _dataSource, "Description");
            var value = new List<string>();


            for (int i = 1; i <= dimensionArray.GetLength(0); i++)
            {
                if (((string)dimensionArray[i, 3].ToUpper()).Equals("COLUMNS"))
                {
                    value.Add((string)dimensionArray[i, 2]);
                }

            }
            return value;
        }

        public List<string> GetRows()
        {
            var dimensionArray = xlApp.Run("SAPListOfDimensions", _dataSource, "Description");
            var value = new List<string>();


            for (int i = 1; i <= dimensionArray.GetLength(0); i++)
            {
                if (((string)dimensionArray[i, 3].ToUpper()).Equals("ROWS"))
                {
                    value.Add((string)dimensionArray[i, 2]);
                }
            }
            return value;
        }

        public int PauseRefresh()
        {
            return (int)xlApp.Application.Run("SAPSetRefreshBehaviour", "Off");
        }

        public int ActivateRefresh()
        {
            return (int)xlApp.Application.Run("SAPSetRefreshBehaviour", "On");
        }

        public int PauseVariableSubmit()
        {
            return (int)xlApp.Application.Run("SAPExecuteCommand", "PauseVariableSubmit", "On");
        }
        public int ActivateVariableSubmit()
        {
            return (int)xlApp.Application.Run("SAPExecuteCommand", "PauseVariableSubmit", "Off");
        }

        public int SetFilter(string name, string value)
        {
            var techName = GetDimensionTechName(name);
            return (int)xlApp.Application.Run("SAPSetFilter", _dataSource, techName, value, "INPUT_STRING");
        }

        public int SetFilters(Dictionary<string, string> filters)
        {
            PauseRefresh();
            int res = 0;
            foreach (var pair in filters)
            {
                res = SetFilter(pair.Key, pair.Value);
                if (res == 0)
                {
                    return res;
                }
            }
            ActivateRefresh();
            return res;
        }

        public string GetLastError()
        {
            return xlApp.Application.Run("SapGetProperty", "LastError", "Text");
        }

        public int AddToRows(string name)
        {
            var techName = GetDimensionTechName(name);
            return (int)xlApp.Run("SAPMoveDimension", _dataSource, techName, "ROWS");
        }
        public int AddToColumns(string name)
        {
            var techName = GetDimensionTechName(name);
            return (int)xlApp.Run("SAPMoveDimension", _dataSource, techName, "COLUMNS");
        }
        public int AddToFilters(string name)
        {
            var techName = GetDimensionTechName(name);
            return (int)xlApp.Run("SAPMoveDimension", _dataSource, techName, "FILTER");
        }
    }

}

