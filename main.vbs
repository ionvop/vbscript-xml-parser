strXMLExample = "<Hello_World><hello><hello>1</hello><world>2</world></hello><world><hello>3</hello><world>4</world></world><hello><hello>5</hello><world>6</world></hello><world><hello>7</hello><world>8</world></world>"

function funcMidString(fInput, fLeft, fRight)
    if instr(fInput, fLeft) or instr(fInput, fRight) then
    else
        funcMidString = false
        exit function
    end if

    funcMidString = mid(fInput, instr(fInput, fLeft) + len(fLeft))

    if instr(funcMidString, fRight) then
    else
        funcMidString = false
        exit function
    end if

    funcMidString = left(funcMidString, instr(funcMidString, fRight) - 1)
end function

function funcCountSubstrings(fInput, fFind, fCase)
    dim arrTemp

    if fCase then
    else
        fInput = lcase(fInput)
        fFind = LCase(fFind)
    end if

    arrTemp = split(fInput, fFind)
    funcCountSubstrings = ubound(arrTemp)
end function

function funcXMLReader(fInput, fOpen, fNumber)
    dim strFirstIndex, strTemp, varStart, strFirstIndexContent, varPosition
    varStart = 1
    fInput = replace(fInput, chr(9), "")

    do while instr(fInput, fOpen)
        if funcMidString(fInput, "<", ">") = fOpen then
            if fNumber > 0 then
                fNumber = fNumber - 1
            else
                if instr(fInput, "</" & fOpen & ">") then
                    varPosition = 1

                    do
                        varPosition = instr(varPosition + 1, fInput, "</" & fOpen & ">")
                        strFirstIndexContent = left(fInput, varPosition - 1)

                        if funcCountSubstrings(strFirstIndexContent, "<" & fOpen & ">", true) - 1 > funcCountSubstrings(strFirstIndexContent, "</" & fOpen & ">", true) then
                        else
                            exit do
                        end if

                        if instr(varPosition + 1, fInput, "</" & fOpen & ">") then
                        else
                            funcXMLReader = false
                            exit function
                        end if
                    loop

                    funcXMLReader = left(fInput, varPosition - 1)
                    funcXMLReader = mid(funcXMLReader, instr(funcXMLReader, "<" & fOpen & ">") + len("<" & fOpen & ">"))
                    exit function
                else
                    funcXMLReader = mid(funcXMLReader, instr(funcXMLReader, "<" & fOpen & ">") + len("<" & fOpen & ">"))
                    exit function
                end if
            end if
        end if

        strFirstIndex = funcMidString(fInput, "<", ">")

        if instr(fInput, "</" & strFirstIndex & ">") then
            varPosition = 1

            do
                varPosition = instr(varPosition + 1, fInput, "</" & strFirstIndex & ">")
                strFirstIndexContent = left(fInput, varPosition - 1)

                if funcCountSubstrings(strFirstIndexContent, "<" & strFirstIndex & ">", true) - 1 > funcCountSubstrings(strFirstIndexContent, "</" & strFirstIndex & ">", true) then
                else
                    exit do
                end if

                if instr(varPosition + 1, fInput, "</" & strFirstIndex & ">") then
                else
                    funcXMLReader = false
                    exit function
                end if
            loop

            fInput = mid(fInput, instr(varPosition, fInput, "</" & strFirstIndex & ">") + len("</" & strFirstIndex & ">"))
        else
            fInput = mid(fInput, instr(fInput, "<" & strFirstIndex & ">") + len("<" & strFirstIndex & ">"))
        end if
    loop

    funcXMLReader = false
end function

function funcXMLDirectory(fXML, fPath)
    dim arrIndex, strIndex
    arrIndex = split(fPath, ",")

    for each strIndex in arrIndex
        fXML = funcXMLReader(fXML, left(strIndex, instr(strIndex, ":") - 1), mid(strIndex, instr(strIndex, ":") + 1))
    next

    funcXMLDirectory = fXML
end function

wscript.echo(funcXMLDirectory(strXMLExample, "hello:1,world:0"))