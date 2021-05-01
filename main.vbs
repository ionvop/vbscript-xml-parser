strXMLExample = "<Hello_World><hello><hello><hello>1</hello><world>2</world><hello>3</hello><world>4</world></hello><world><hello>5</hello><world>6</world><hello>7</hello><world>8</world></world><hello><hello>9</hello><world>10</world><hello>11</hello><world>12</world></hello><world><hello>13</hello><world>14</world><hello>15</hello><world>16</world></world></hello><world><hello><hello>17</hello><world>18</world><hello>19</hello><world>20</world></hello><world><hello>21</hello><world>22</world><hello>23</hello><world>24</world></world><hello><hello>25</hello><world>26</world><hello>27</hello><world>28</world></hello><world><hello>29</hello><world>30</world><hello>31</hello><world>32</world></world></world><hello><hello><hello>33</hello><world>34</world><hello>35</hello><world>36</world></hello><world><hello>37</hello><world>38</world><hello>39</hello><world>40</world></world><hello><hello>41</hello><world>42</world><hello>43</hello><world>44</world></hello><world><hello>45</hello><world>46</world><hello>47</hello><world>48</world></world></hello><world><hello><hello>49</hello><world>50</world><hello>51</hello><world>52</world></hello><world><hello>53</hello><world>54</world><hello>55</hello><world>56</world></world><hello><hello>57</hello><world>58</world><hello>59</hello><world>60</world></hello><world><hello>61</hello><world>62</world><hello>63</hello><world>64</world></world></world>"

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
    fInput = replace(fInput, vbCrlf, "")

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

'What's 9 + 10?
wscript.echo(funcXMLDirectory(strXMLExample, "world:0,world:0,hello:0")) 'Expected result: 21 (Success)
