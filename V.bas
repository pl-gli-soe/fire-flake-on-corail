Attribute VB_Name = "V"
'The MIT License (MIT)
'
' Copyright (c) 2019 FORREST
' Mateusz Milewski mateusz.milewski@opel.com aka FORREST
'
' The QT - quickTool
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.


' name of this software  due to fact that the main logic was written in a couple of days :P


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2017-11-13
' v0.1 init on this project
' 3 cfg sheets: input, register, plt-list
' OOP schema ICorail -> Corail Blue & Orange - a plan
' also plan to have app.run (kind of multi-thread app)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2017-11-14
' v0.2 next steps with new classes:
' parser
' rawdata
' shellhandler
' eventhandler connected with corail handler
' sets of corails
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2017-11-16
' dopisanie implemenacji odpowiedzialnej za frame:
' Set .frame = .doc.frames(FFOC.G_MAIN_FRAME_ID)
' okazalo sie ze orange corail jest strona w stronie - musialem to jakos obejsc...
'
' new class: DropperHandler
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2017-11-20
'v0.4
'duzo zmian
'lacznie z pierwszym udanym polaczeniem z danymi na zywym systemie
'jest to pierwsza podwersja pisana bezposrednio na francuskim sprzecie
'testy natychmiastowe bez koniecznosci przeklikiwania sie pomiedzy mailami
' poprawiony parser
' ujednolicone dzialania pomiedzy corailami blue and orange
' schema:
'CorailHelper -> CorailRunner -> ICorail jako interfejs - orange oraz blue korzystaja z tych samych metod

' Orange, Blue, Manual Corail implements ICorail

''w manual Corail wszystkie metody wlasciwie wygldaja tak samo jak w interfejsie - spowodowane jest to glownie brakiem danych pobiernaych
' wiec generalnie jest pusto i cicho - jedyna zmiana to zaprzestanie wyrzucania bledow krytycznych jesli pod koniec logiki dane wciaz
' sa nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2017-11-21
'v0.5
' nowe funkcje:
' 1 open plants
' 2 close all corails and maestros
' 3 after open plants ie is not visible
' 4 initial layout for the tool
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2017-11-22
'v0.6
' waiting for IE not working need ta adjust more directly with content of corail site
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-03-06
' v0.7
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' adjust for safe mode in IE
' removal of some logic inside layout changes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



' 2018-03-29
' v0.8
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' add export this project module for githib repository...
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



' 2018-04-09
' v0.9
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' fix on datadownload - some issues behind taking data and wrong count on balance taking zero from decimal places
' as a "normal" zeros - to be fix on this version
' + dropper handler - added backlog ficzer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-04-10
' v0.91
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' temp solution for maestro will be treat as manual plan - only filled by zeros and formulas
' some extra fixes on dates and issues on out of range possibility also to be fixed in near future with
' ranges which are too long - some limitation required from end-user.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-04-11
' v0.92
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' initial implementation for Maestro
' still errors on multi order and requirements numbers - if red font then showing zero - to be fix on 0.93
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



' 2018-04-12
' v0.93
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' changes on layout - be more like fire flake
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-04-16
' v0.94
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' unfinished fire flake layout
' fix in orange requirements data download - to test! - check implementation
' change on parser:
' pCmdCatcher -> pCmdCatcher1 + pCmdCatcher2
' and
' pExpCatcher -> pExpCatcher1 + pCmdCarcher2
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-04-16 II
' v0.95
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' layout more like fire flake and fill rest of the common data
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-04-16 III
' v0.01
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' skip to new name -> QT starts to be FF
' to fix;
' no colors on stock
' no filter
' no freeze
' first runout without runout after sorting on top, which is no so right and fine
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



' 2018-04-16 IV
' v0.02
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' skip to new name -> QT starts to be FF
' to fix;
' no colors on stock
' no filter
' no freeze
' first runout without runout after sorting on top, which is no so right and fine
' new addons - input comments - converted plt names
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-04-23
' v0.03
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' fix on part numbers which do not have data in corail or maestro
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



' 2018-04-24
' v0.04
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' additional fix on stock
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-06-14
' v0.05
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' new approach on dealing with ending balance on corail on yesterdays dates (split stock on stock depart and reception)
' try to remove connection on format od the date...
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-06-25
' v0.06
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' safe mode for version 0.05
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 2018-06-25
' v0.07
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' final fixes - layout adjustment - prototype for tests.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




' v0.08
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' fix on dropper for manual parts...
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' v0.11
' foo loaded in IEHandler adjusted



' v0.20
' try to use winHttp


' v.21
' added part name + supplier


' v.22 happy copy during tests for tychy plant - performance check on more than 100 parts

' v.23 - monior fixes on layout

' v.24 login - label change - preperation for hourly and weekly


' v.25
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'fix for decimal places in coverage + minor fixes with decimal treatment
'on parser (important) I've used: application.DecimalSeparator
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' v.27
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' feature - added hazards into stock if checkbox value is true
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



' v.29
' 2019-09-17
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' ParserHandler change new class name in order and ship for sau
'Order and ship quantity have different ecwkeyname :
'SX order example: <td ecwkeyname="orderedQuantity">
'<div class="ecwButtonTexteOverOrdered number ecwButtonTexteOver">0.0</div></td>
'LU order example: <td ecwkeyname="sauOrderedQuantity">
' <div class="number ecwButtonTexteOverSauOrdered ecwButtonTexteOver">0.0</div></td> // but this "sau" is optional!
'
'
'SX ship example:
' <td ecwkeyname="orderedQuantity"><div class="ecwButtonTexteOverOrdered number ecwButtonTexteOver">0.0</div></td>
'LU ship example:
' <td ecwkeyname="sauOrderedQuantity"><div class="number ecwButtonTexteOverSauOrdered ecwButtonTexteOver">0.0</div></td> // but this "sau" is optional!

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' v.30
' 2019-10-04
' MAESTRO STOCK FIX!!!



' v.31
' fix on open on chrome from corail parts using shell
' my custom logic for decimal point reference from register Q17


' v.32
' some fix with BIP param

' v.33
' initial preparation for data from 2510

' v.34
' some stuff in plt list

' v.35
' adjust input for two chars

' v.5
' duza zmiana - wrzucamy 2510 do implementacji!
' plus dobre praktyki z wcoca

' v.51
' poprawka pod 2510 - jesli danych nie - to trzeba to ochronic


' v.52
' jednak milestony w koment + godzina potrzebna


' v.53
' wymuszanie formatu yyyy-mm-dd hh:mm

' V 0.54
' dropper 170
'r.Offset(0, 1).AddComment txtForCmmnt
'r.Offset(0, 1).Comment.Shape.TextFrame.AutoSize = True
'r.Offset(0, 1).Comment.Shape.TextFrame.Characters.Font.Name = "Courier New"
'r.Offset(0, 1).Comment.Shape.TextFrame.Characters.Font.Size = 8

'v.055
' extra spaces!


'v.0.56
' additional type checking in Parser class - there was some mismatches.

'v0.57
' change in class CorailItem2510 for better pasing double types


' v.60
' problem with hazards... added extra column (M)


' v61
' this version prohibits possibility of putting hazards into formula for ending balance.


' v65
' skip to 65 to see diff - first version with make list feature!
' pre-list on place as helper to define regular input list
' Make1 as main form for maintain input data initially


' v66
' issue with new version of the corail
' data are not taken propetly with the macro.
'
' new feature : iterationOfgetData - 5 times check http request if response is not OK
' in class HTTP Request Handler.

' comments with transport details not appearing! NOK NOK NOK


'v67
' optiona BO as string in ICorail interface - bo stands for B-O columns which are just common data - means i want to have just lightweight info without coverage - new special button
' errors in comments


' v68 and 69
' this is major issue to fix - somehow data are not downloaded correct in a few exmaples
' so solution was simple: iterationOfgetData - was taking loop if http post was not working properly and by force and i forgot pass args param!


'v70 krytyczna zmiana routingu
' narazie tylko VH


' v071
' prototype with static recognition on plt-list for new protocal of OAuth

' v072
' KA issue - timeout issue -> re-check implementation of recursion if request was not "valid"

' v073 -> fix rawstring => .rawstring
'
'        If POST_FLAG = "POST" Then
'
'            If tmpStr1 Like "*indicePage=*pageSize=*numberOfElement=*" Then
'                iterationOfgetData = CStr(odp)
'            Else
'                iterationOfgetData = "" ' tryToGetDataAgain(odp, times - 1, url, login, pass, POST_FLAG, args)
'            End If
'
'        Else
'v074 -> enable retriving data from Corail without data from 2510

'        Added in GlobalModule :
'           Global IS2510REQ As Boolean
'
'        Added in HttpRequestHeader class:
'           If e = BLUE And GlobalModule.IS2510REQ = True Then
'
'        Added in Login Form
'           New Button "Run Fire Flake Without 2510"
'
'v075 -> extended horizon
'enable retiriving data from extended coverage, thanks to generating another request to 2720

'Functions that have been added:
'getExtendedCoverage()
'getRequestedDomExtendedHorizon()
'getcollectonOfXtraCorailItems()

'Classes added:
'Corail, Corail_2720_Screen, CorailData, CorailDataFrom2720, CoverageItem, Dh_Handler, ICorailData, InputConfigHandler,
'IOrderDetails, OrderItem, Parser, SuitableData2720, SuitableXtraData2720, Validator, WeeklyCoverage, XtraRqmItem, IPArser


' to be 076 - error on parseQtyFrom2510 -var tmpstr still not having valid number
' to be 076 - new in Parser (WCOC) and ParserHandler (FF) : Private Function focusOnSupplierName(sn As String) As String
' to be resolved in 077 - remove classes from Weekly Coverage - 076 implementation is against DRY rule.

' to be 077 - changes on 2510 logic - from now on </div and </DIV are different things - be careful - library providing DOM object are now case sensitive
' to be 077 - test required on different machines - to be clarified
' ParserHandler instance - issue on case sensitive : Private Function removeDivs(strWithDiv As String) As String - update on this private function




'to be 078 - removing all classes that have been implemented from Weekly Coverage in 0.75
'classes deleted: 'Corail, Corail_2720_Screen, CorailData, CorailDataFrom2720, CoverageItem, Dh_Handler, ICorailData, InputConfigHandler,
'IOrderDetails, OrderItem, Parser, SuitableData2720, SuitableXtraData2720, Validator, WeeklyCoverage, XtraRqmItem, IPArser

'removing data repackaging process: Weekly Coverage objects -> Fire flake objects; Creating directly New CorailItem(FF) instead of CoverageItem(WC)
'ParserHandler:Ln 1794, 1800, 1874

'Adding Weekly Coverage's atrributes to CorailItem:
'Public sgrLine As String
'Public clv As Double
'Public fabPlan As Double
'Public Total As Double

'Fire Flake 0.78 -> data flow
'CorailRunner.zdarzenie_initCorail()
'   new CorailBlue()
'       Parser=new ParserHanlder() //first parser
'   .generateInnerHttpRequest()
'       new HttpRequestHandler()
'       .init()
'       .braceWithDom()
            '   .getData() // FireFlake request 2720
            '   .getExtendedCovData()// WeeklyCoverage request 2720
            '   Set theParser = New ParserHandler //second Parser
                        
            '   theParser.importPackageOfData dom.extdoc //putting WeeklyCoverage as HTMLDocument as parameter
                    '.innerParse2720() // parsing data from weekly coverage htmldocument
                        'Set collectonOfXtraCorailItems = New CorailIteration
                            'Adding to collectonOfXtraCorailItems CorailItem
            '   Set collectonOfXtraCorailItems = theParser.getConvertedDataSuitableForExcel()
            
            'Set collectonOfXtraCorailItems = theParser.getConvertedDataSuitableForExcel() // returns CorailIteration of CorailItems from Weekly Cov.
            
            'Do While Loop
                'Set dom2510 = New DOM2510Handler
                'dom2510.httppost() //getting data from 2510 for Fire Flake
                
        'Set collectonOfXtraCorailItems = req.getCollectonOfXtraCorailItems() //collectonOfXtraCorailItems As CorailIteration
        'bringing the collection of CorailItems Form Weekly Coverage
        
    'Set dane = .getData() //CorailBlue method
        'Function getDate()
            'Parser.htmlDataIntoCovertedData(collectonOfXtraCorailItems AS CorailIteration) //invoked on first parser, use CorailIteration set 5 lines earlier
                'collectonOfXtraCorailItems=collectonOfXtraCorailItemsBrangToMerge // corailIteration of CorailItems from Weekly Coverage
                '.htmlTableToRawMatrix(collectonOfXtraCorailItemsBrangToMerge AS CorailIteration)
                    'Set tmp = New ConvertedData //setting wrapper for CorailIteration
                    'Set ii = New CorailIteration
                        'parsing HTMLTable that comes in htmlTableToRawMatrix as parameter
                        'i=new CorailItem()
                    'ii.add i //creating CorailIteration of CorailItems from FireFlake
                    
                    'Set filteredCollectonOfXtraCorailItemsBrangToMerge = New Collection //creating collection for filtered data by sgrLine
                    'filtering data
                    'Reset of collectonOfXtraCorailItemsBrangToMerge.pItems to filteredCollectonOfXtraCorailItemsBrangToMerge
                    
                    'For loop that goes for each CorailItem in collectonOfXtraCorailItemsBrangToMerge.pItems and adds it to ii(CorailIteration of
                    'Fire Flake CorailItems), only if date of each Weekly Cov. CorailItem is later than last date of Fire Flake CorailItem.
                    'At the end we get a ii CorailIteration with all data.
                    
