# excel-auto-email

This macro writes and sends emails, given a list of names and email addresses.

The aim was for this to be used as a template.  For example, emails are sent on the condition that the entry in the 'On Mailing List' column is "Yes".  This condition can be easily removed or modified.

A sample of the dataset for which this macro was designed is pictured below.  The list of names begins at cell B3, the 'Email Address' column is horizontally offset by 1, the 'On Mailing List' column is offset by a further 1, and the administrator email (the address from which emails are sent, and to which confirmation is sent) is specified in cell G7.  All of these can be changed, so long as their references are updated in the VBA file.

Note that for development purposes, emails are set to display (with the option to send) instead of send automatically.  This can be changed easily by replacing 'EmailItem.Display' with 'EmailItem.Send'.

Here is a sample of the dataset for which this macro was designed:

<img src="https://github.com/hllewellyn1/excel-auto-email/blob/09edc0753c967af19bb40ca2f4aad6a1d0adb3bf/sample_data.png" width="500">

##
[![Generic badge](https://img.shields.io/badge/VERSION-1.0-<COLOR>.svg)](https://github.com/hllewellyn1/HaskAD)
