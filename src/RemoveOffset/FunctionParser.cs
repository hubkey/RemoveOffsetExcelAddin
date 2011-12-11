using System;
using System.Collections.Generic;
using System.Text;

namespace RemoveOffset
{
    public enum ParserState
    {
        None,
        Parsing,
        Error,
        Success
    }

    public class ParserResult
    {
        public int Pos { get; private set; }
        public string Value { get; private set; }

        public ParserResult(int pos, string value)
        {
            Pos = pos;
            Value = value;
        }
    }

    public class FunctionParser
    {
        private const char OpenChar = '(';
        private const char CloseChar = ')';
        private readonly string _functionName;
        private readonly Func<ParserResult, string> _substitutionFunction;
        private readonly int _length;
        private string _input;
        private ParserState _state = ParserState.Error;
        private int _pos;
        private int _headPos;

        public FunctionParser(string functionName, string input)
            : this(functionName, input, null)
        {
        }

        public FunctionParser(string functionName, string input, Func<ParserResult, string> substitutionFunction)
        {
            _substitutionFunction = substitutionFunction;
            _functionName = functionName;
            _input = input;
            _length = _input.Length;
        }

        public List<ParserResult> Results { get; private set; }
        public string Output { get { return _input; } }

        public ParserState State
        {
            get { return _state; }
        }

        public bool Parse()
        {
            _state = ParserState.Parsing;
            Results = new List<ParserResult>();
            try
            {
                while (ReadFunctionName())
                {
                    var result = new ParserResult(_headPos, string.Concat(_functionName, ReadFunctionBody()));
                    Results.Add(result);
                    if (_substitutionFunction == null) continue;
                    var newValue = _substitutionFunction(result);
                    _input = _input.Remove(result.Pos, result.Value.Length);
                    _input = _input.Insert(result.Pos, newValue);
                    _pos = result.Pos + newValue.Length;
                }

            }
            catch
            {
                _state = ParserState.Error;
                return false;
            }
            _state = ParserState.Success;
            return true;
        }

        private string ReadFunctionBody()
        {
            var body = new StringBuilder();
            var open = 0;
            while (!EOF())
            {
                body.Append(_input[_pos]);
                switch (_input[_pos++])
                {
                    case OpenChar:
                        ++open;
                        break;
                    case CloseChar:
                        --open;
                        break;
                }
                if (open == 0)
                    break;
            }
            return body.ToString();
        }

        private bool EOF()
        {
            return _pos < 0 || _pos >= _length || _headPos < 0 || _headPos > _length;
        }
        private bool ReadFunctionName()
        {
            if (EOF())
                return false;
            _headPos = _input.IndexOf(_functionName + OpenChar, _pos);
            if (!EOF())
                _pos = _headPos + _functionName.Length;
            return !EOF();
        }
    }
}
