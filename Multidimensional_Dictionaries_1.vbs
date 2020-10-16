Set zoo = CreateObject("Scripting.Dictionary")

allMammals = Array("Donkey", "Horse", "Goat")
allReptiles = Array("Lizard", "Turtle", "Snake")

zoo.Add "Mammals", Join(allMammals)
zoo.Add "Reptiles", Join(allReptiles)

MsgBox Join(zoo.Keys()) 'return the key values of the two dictionaries
MsgBox zoo.Item("Mammals")
MsgBox Join(zoo.Items())