def add_header(string):
    string +="   <HEADER>\n"
    string +="       <TALLYREQUEST>Import Data</TALLYREQUEST>\n"
    string +="   </HEADER>\n"
    return string

def add_bodyStarting(string):
    string += "   <BODY>\n"
    return string

def add_bodyEnding(string):
    string += "   </BODY>\n"
    return string