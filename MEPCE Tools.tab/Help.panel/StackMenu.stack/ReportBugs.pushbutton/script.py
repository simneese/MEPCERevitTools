# -*- coding: utf-8 -*-
__title__ = "Report Bugs"
#__highlight__ = "new"
__doc__ = """Version = 1.0
Date    = 2025.06.16
_________________________________________________________________
Description:
This button will open a bug report form.
_________________________________________________________________
How-to:
-> Click button
-> Fill out bug report
-> Send report to sneese@mepce.com
_________________________________________________________________
Last update:
- [2025.06.16] - 1.0 RELEASE
_________________________________________________________________
Author: Simeon Neese"""

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from System import DateTime
from System.Windows import Window, Thickness, HorizontalAlignment, VerticalAlignment, MessageBox, MessageBoxButton, MessageBoxImage, WindowStartupLocation
from System.Windows.Controls import StackPanel, TextBox, Button, TextBlock, ScrollBarVisibility
from System.IO import File, Directory
import os

# Safe fallback for TextWrapping.Wrap
TextWrapping = getattr(TextBox(), 'TextWrapping')

class ReportBugWindow(Window):
    def __init__(self):
        self.Title = "Report a Bug"
        self.Width = 450
        self.Height = 500
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = 0  # NoResize

        panel = StackPanel()
        panel.Margin = Thickness(10)

        # Bug Title
        panel.Children.Add(self._make_label("Bug Title:"))
        self.textbox_title = TextBox()
        self.textbox_title.Margin = Thickness(0, 0, 0, 10)
        panel.Children.Add(self.textbox_title)

        # Bug Description
        panel.Children.Add(self._make_label("Description:"))
        self.textbox_desc = TextBox()
        self.textbox_desc.TextWrapping = TextWrapping.Wrap
        self.textbox_desc.AcceptsReturn = True
        self.textbox_desc.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        self.textbox_desc.Height = 100
        self.textbox_desc.Margin = Thickness(0, 0, 0, 10)
        panel.Children.Add(self.textbox_desc)

        # Error Message
        panel.Children.Add(self._make_label("Paste Error Message:"))
        self.textbox_error = TextBox()
        self.textbox_error.TextWrapping = TextWrapping.Wrap
        self.textbox_error.AcceptsReturn = True
        self.textbox_error.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        self.textbox_error.Height = 100
        self.textbox_error.Margin = Thickness(0, 0, 0, 10)
        panel.Children.Add(self.textbox_error)

        # Email Instructions
        email_instr = TextBlock()
        email_instr.Text = "After submitting, please email this report to: sneese@mepce.com"
        email_instr.Margin = Thickness(0, 0, 0, 10)
        panel.Children.Add(email_instr)

        # Submit Button
        btn_send = Button()
        btn_send.Content = "Submit"
        btn_send.Width = 100
        btn_send.HorizontalAlignment = HorizontalAlignment.Right
        btn_send.Click += self.on_send_click
        panel.Children.Add(btn_send)

        self.Content = panel

    def _make_label(self, text):
        label = TextBlock()
        label.Text = text
        label.Margin = Thickness(0, 0, 0, 5)
        return label

    def on_send_click(self, sender, event):
        title = self.textbox_title.Text.strip()
        desc = self.textbox_desc.Text.strip()
        error = self.textbox_error.Text.strip()

        if not title or not desc:
            MessageBox.Show(
                "Please enter both a title and description.",
                "Missing Info",
                MessageBoxButton.OK,
                MessageBoxImage.Warning
            )
            return

        timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss")
        filename = "BugReport_{}.txt".format(timestamp)
        folder = os.path.expanduser("~\\Desktop")

        if not os.path.exists(folder):
            os.makedirs(folder)

        full_path = os.path.join(folder, filename)

        content = u"BUG TITLE:\n{}\n\nTIMESTAMP:\n{}\n\nDESCRIPTION:\n{}\n\nERROR MESSAGE:\n{}".format(title, timestamp, desc, error)
        File.WriteAllText(full_path, content)

        MessageBox.Show(
            "Bug report saved to:\n{}\n\nPlease email the file to: sneese@mepce.com".format(full_path),
            "Report Saved",
            MessageBoxButton.OK,
            MessageBoxImage.Information
        )

        self.Close()

# Run the window
if __name__ == "__main__":
    win = ReportBugWindow()
    win.ShowDialog()
