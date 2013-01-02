using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DNA.Common
{
    public struct Column
    {
        private int _column;

        public Column(int initalValue)
        {
            if (initalValue < 1)
            {
                throw new ArgumentException("initialValue");
            }
            _column = initalValue;
        }

        public Column(string initialValue)
        {
            if (string.IsNullOrWhiteSpace(initialValue))
            {
                throw new ArgumentException("initialValue");
            }
            int order = initialValue.Length;
            int value = 0;
            for (int i = 0; i < order; i++)
            {
                char c = initialValue[i];
                if (c < 65 || c > 90)
                {
                    throw new ArgumentException("initialValue");
                }
                int place = (int)Math.Pow(26.0, order - i - 1);
                int factor = c - 64;
                value += factor * place;
            }

            _column = value;
        }

        /*public static bool operator <(Column a, int b)
        {
            return a._column < b;
        }

        public static bool operator <(Column a, int b)
        {
            return a._column < b;
        }

        public static bool operator >(Column a, int b)
        {
            return a._column > b;
        }*/

        public static implicit operator int(Column c)
        {
            return c._column;
        }

        public static implicit operator Column(int i)
        {
            return new Column(i);
        }

        /*public static Column operator +(Column a, int b)
        {
            if (a._column - b > 0)
            {
                a._column += b;
                return a;
            }
            throw new ArithmeticException();
        }

        public static Column operator +(int b, Column a)
        {
            return a + b;
        }

        public static Column operator ++ (Column input)
        {
            input._column++;
            return input;
        }

        public static Column operator --(Column input)
        {
            input._column--;
            return input;
        }*/

        public override string ToString()
        {
            
            /*int order = (_column == 0) ? 1 : (int)Math.Floor(Math.Log(_column, 27.0)) + 1;
            char[] characters = new char[order];
            int dividend = _column;
            for (int i = 0; i < order; i++)
            {
                int place = (int)Math.Round(Math.Pow(27.0, order - i - 1));
                int val = dividend / place;
                characters[i] = (char)(val + 64);
                dividend = _column - val * place + 1;
            }*/


            int dividend = _column;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName.ToString();
        }
    }
}
