﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace MayaConnected
{
    class CategoryAttribute : TraitAttribute
    {
        public CategoryAttribute(string categoryName)
            : base("Category", categoryName)
        {
        }
    }
}
