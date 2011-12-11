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
        private readonly string _input;
        private string _output;
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
            _output = input;
            _length = _output.Length;
        }

        public List<ParserResult> Results { get; private set; }
        public string Input { get { return _input; } }
        public string Output { get { return _output; } }

        public ParserState State
        {
            get { return _state; }
        }

        public ParserState Parse()
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
                    _output = _output.Remove(result.Pos, result.Value.Length);
                    _output = _output.Insert(result.Pos, newValue);
                    _pos = result.Pos + newValue.Length;
                }

            }
            catch
            {
                return _state = ParserState.Error;
            }
            return _state = ParserState.Success;
        }

        private string ReadFunctionBody()
        {
            var body = new StringBuilder();
            var open = 0;
            while (!EOL())
            {
                body.Append(_output[_pos]);
                switch (_output[_pos++])
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

        private bool EOL()
        {
            return _pos < 0 || _pos >= _length || _headPos < 0 || _headPos > _length;
        }
        private bool ReadFunctionName()
        {
            if (EOL())
                return false;
            _headPos = _output.IndexOf(_functionName + OpenChar, _pos);
            if (!EOL())
                _pos = _headPos + _functionName.Length;
            return !EOL();
        }
    }
}
