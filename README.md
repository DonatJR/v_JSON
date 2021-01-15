# v_JSON
A JSON parser for VBScript.

Use the main_template like so:

    Dim json
    Set json = New v_JSON

    json.FromString "{""key1"": null, ""key2"": { ""key3"": ""val3"" }, " & _
                    """key4"": ""val4"", ""key5"": true, ""key6"": 7.8, " & _
                    """employees"":[ { ""firstName"":""John"", ""lastName""" & _
                    ":""Doe"" }, { ""firstName"":""Anna"", ""lastName"":" & _
                    """Smith"" }, { ""firstName"":""Peter"", ""lastName"":" & _
                    """Jones"" } ] }"
                    
    WScript.Echo json.Item("employees")(2).Item("firstName")

This will produce the following output in the console or MsgBox:

    Peter
