﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HermesLabelCreator.Models
{
    public class SpreadsheetCellDto
    {
        public string CellColumnName { get; internal set; }
        public string Value { get; internal set; }
    }
}
