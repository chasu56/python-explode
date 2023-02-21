# python-explode

If you have a excel in following format
Name   Age Nationality  Sex
Alex, Ram, Harry 23, 24,25 USA, UK, India M,F,M

This code explodes to create one sheet with following format
Name   Age Nationality  Sex
Alex 23 USA M
Ram 24 UK F
 Harry 25 India M 
If exploded value count do not match, it puts '' as value 
Output formatted structure to a different excel book
