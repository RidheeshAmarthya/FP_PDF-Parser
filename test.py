import re

teststr = "4E+343434"

print(re.findall(".*\dE\+|\-\d.*",teststr))