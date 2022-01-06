using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OslSpreadsheet.Models
{
    public static class oSpreadsheetExtensions
    {
        /// <summary>
        /// Convert cell to type of float
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static oCell AsFloat<T>(this oCell cell, float value)
        {
            var c = cell;
            c.Value = value.ToString();
            c.ValueType = CellValueType.Float;

            return c;
        }
    }
}
