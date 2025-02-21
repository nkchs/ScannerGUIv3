using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.UI.Xaml.Controls;

namespace ScannerGUIv3.Services;

public static class ConsoleService
{
    private static TextBox? _consoleOutput;

    public static void Initialize(TextBox consoleOutput)
    {
        _consoleOutput = consoleOutput;
    }

    public static void WriteLine(string text)
    {
        if (_consoleOutput != null)
        {
            //_consoleOutput.Text += text + Environment.NewLine;
            _consoleOutput.Text = text + Environment.NewLine + _consoleOutput.Text;
            _consoleOutput.SelectionStart = _consoleOutput.Text.Length;
            _consoleOutput.SelectionLength = 0;
        }
    }
}
