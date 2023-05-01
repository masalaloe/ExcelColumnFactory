using HRASystem.Libraries.ExcelUtilities.Builder;
using HRASystem.Libraries.ExcelUtilities.Factory;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HRASystem.Libraries.ExcelUtilities
{
    public abstract class OptionsOwnerBuilder
    {
        internal IDictionary<string, object> Options { get; private set; } = new Dictionary<string, object>();

        //protected internal void Collection<T>(string optionName, Action<ColumnFactory> configurator)
        //{

        //}
    }
}

namespace HRASystem.Libraries.ExcelUtilities.Factory
{
    public class ColumnsFactory 
    {
        public ColumnsFactory() { }

        public ColumnsFactoryBuilder Builder() => new ColumnsFactoryBuilder();
    }

    public class ColumnsFactoryBuilder
    {
        private List<Action<ColumnBuilder>> columnBuilderActions = new List<Action<ColumnBuilder>>();

        public ColumnsFactoryBuilder AddColumn(Action<ColumnBuilder> configurator)
        {
            columnBuilderActions.Add(configurator);
            return this;
        }

        public List<IDictionary<string, object>> Build()
        {
            var r = new List<IDictionary<string, object>>();
            foreach (var configurator in columnBuilderActions)
            {
                var columnBuilder = new ColumnBuilder();
                configurator.Invoke(columnBuilder);
                r.Add(columnBuilder.Options);
            }
            return r;
        }
    }
}

namespace HRASystem.Libraries.ExcelUtilities.Builder
{
    public class ColumnBuilder : OptionsOwnerBuilder
    {
        private void DefaultValue()
        {
            base.Options["width"] = 12;
        }

        public ColumnBuilder() 
        {
            DefaultValue();    
        }

        public ColumnBuilder Width(int value)
        {
            base.Options["width"] = value;
            return this;
        }

        public ColumnBuilder Format(string value)
        {
            base.Options["format"] = value;
            return this;
        }

        public ColumnBuilder Validation(Action<ValidationBuilder> configurator)
        {
            var f = new ValidationBuilder();
            configurator.Invoke(f);
            base.Options["validation"] = f.Options;
            return this;
        }

        public ColumnBuilder Header(Action<HeaderBuilder> configurator)
        {
            var f = new HeaderBuilder();
            configurator.Invoke(f);
            base.Options["header"] = f.Options;
            return this;
        }

        public ColumnBuilder AddHeader(string headerCaption)
        {
            var f = new HeaderBuilder();
            var g = new FontBuilder();
            f.Options["caption"] = headerCaption;
            g.Options["size"] = 12;
            f.Options["font"] = g.Options;
            base.Options["header"] = f.Options;
            return this;
        }

    }

    public class HeaderBuilder : OptionsOwnerBuilder
    {
        public HeaderBuilder Caption(string value)
        {
            base.Options["caption"] = value;
            return this;
        }

        public HeaderBuilder Font(Action<FontBuilder> configurator)
        {
            var f = new FontBuilder();
            configurator(f);
            base.Options["font"] = f.Options;
            return this;
        }

    }

    public class ValidationBuilder : OptionsOwnerBuilder
    {
        public ValidationBuilder ValidationFormula(string value)
        {
            base.Options["validationFormula"] = value;
            return this;
        }

        public ValidationBuilder ValidationFormula(IList<string> value)
        {
            base.Options["validadionFormulaList"] = value;
            return this;
        }
    }

    public class FontBuilder : OptionsOwnerBuilder
    {
        public FontBuilder()
        {
            base.Options["bold"] = true;
        }

        public FontBuilder Size(int value)
        {
            base.Options["size"] = value;
            return this;
        }

        public FontBuilder Bold(bool value)
        {
            base.Options["bold"] = value;
            return this;
        }
    }
}
