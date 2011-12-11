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

    public class Parser
    {
        private const char OpenChar = '(';
        private const char CloseChar = ')';
        private const string Head = "OFFSET";
        private readonly int _length;
        private string _input;
        private ParserState _state = ParserState.Error;
        private int _pos;
        private int _headPos;

        public Parser(string input)
        {
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
            return Parse(null);
        }

        public bool Parse(Func<ParserResult, string> substitutionFunction)
        {
            _state = ParserState.Parsing;
            Results = new List<ParserResult>();
            try
            {
                while (ReadHead())
                {
                    var result = new ParserResult(_headPos, string.Concat(Head, ReadBody()));
                    Results.Add(result);
                    if (substitutionFunction == null) continue;
                    var newValue = substitutionFunction(result);
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

        private string ReadBody()
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
        private bool ReadHead()
        {
            if (EOF())
                return false;
            _headPos = _input.IndexOf(Head + OpenChar, _pos);
            if (!EOF())
                _pos = _headPos + Head.Length;
            return !EOF();
        }
    }
}
