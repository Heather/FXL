namespace RNExcel.Model

open System
open System.IO
open System.Linq

open System.Text.RegularExpressions

open NPOI.HSSF.UserModel

type ExcWor() =
    member X.make(fname,varCol,valCol,startRow,endRow) =
        let cXL name =  
            if name <> "" then
               (name.ToLower().ToCharArray()
                |> Seq.map (fun char -> Convert.ToInt32 char - 96)
                |> Seq.sumBy(fun x -> x + 25)) - 26
            else 0
        if [|fname;varCol;valCol;startRow;endRow|] 
            |> Seq.forall(fun x -> (x <> "" && x <> null)) then 
            using(new FileStream(fname, FileMode.Open, FileAccess.Read))<| fun fs               ->
                let templateWorkbook = new HSSFWorkbook(fs, true)
                let sheet = templateWorkbook.GetSheet("Sheet1")

                let cvar    = cXL varCol   
                let cval    = cXL valCol   

                let IntAndString value =    
                    let (|Match|_|) pattern input =
                        let m = Regex.Match(input, pattern) in
                        if m.Success then Some ([ for g in m.Groups -> g.Value ]) else None
                    match value with
                        | Match @"((?>\d+))(\w+)" x -> Some(x)              
                        | Match @"((?>\d+))" x      -> Some(x @ ["items"])  
                        | Match @"(\w+)" x          -> Some(x)             
                        | _                         -> None

                let sr =
                    try Int32.Parse startRow 
                    with _ -> 0

                match sr with
                | 0 -> sprintf "Error with rows %s" startRow
                | _ ->
                    let doMathAndSave er =
                        let istherewasanerror   = ref false
                        let error               = ref ("", 0) 
                        let writeError error =                
                            let valCol, i = error
                            sprintf "Error getting excel sheet on %s %d" valCol i
                        let vvlist = [ for i in sr..er -> 
                                                            try
                                                               (sheet.GetRow(i-1).GetCell(cvar).ToString(), 
                                                                sheet.GetRow(i-1).GetCell(cval).ToString()) 
                                                            with _ -> 
                                                                istherewasanerror   := true
                                                                error               := (valCol, i)
                                                                "",""  ]

                        if !istherewasanerror then
                            writeError !error   
                        else
                            for i in sr..er do  
                                let varC,valC   = vvlist.Item(i-sr) 
                                let varXCom     = IntAndString varC 
                                if varXCom.Value.Length  = 3 then   
                                    let parsed1 = Double.Parse(valC)
                                    let parsed2 = Double.Parse(varXCom.Value.[1])
                                    sheet.GetRow(i-1).GetCell(cvar).SetCellValue(varXCom.Value.[2])
                                    sheet.GetRow(i-1).GetCell(cval).SetCellValue((parsed1*parsed2))

                            sheet.ForceFormulaRecalculation = true |> ignore 

                            using(new MemoryStream()) <| fun ms ->  
                                templateWorkbook.Write(ms)         
                                let msA = ms.ToArray()
                                using(new FileStream((@"REXCEL.xls"), FileMode.OpenOrCreate , FileAccess.Write))
                                <| fun newF ->
                                    try
                                        newF.Write(msA,0,msA.Length)
                                        sprintf "RExceL.xls created, check the result"
                                    with _ -> "Can't write to file"


                    let er = match endRow with
                                                "0" -> 
                                                        let rec counter cn =
                                                            try
                                                                ignore <| sheet.GetRow(cn).GetCell(cvar)
                                                                ignore <| sheet.GetRow(cn).GetCell(cval)
                                                                counter (cn+1)
                                                            with _ -> (cn-1) 
                                                        counter 0
                                                | _ ->
                                                    try Int32.Parse endRow
                                                    with _ -> 0
                    doMathAndSave er
        else
            "Input paeameters Error"