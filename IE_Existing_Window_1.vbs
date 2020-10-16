Option Explicit

Dim internet_explorer

Set internet_explorer = CreateObject("InternetExplorer.Application")

internet_explorer.Navigate "https://www.youtube.com/results?search_query=python+web+development+career+path"

internet_explorer.Visible = 1