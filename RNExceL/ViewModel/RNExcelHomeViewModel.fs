namespace RNExcel.ViewModel

open System
open System.Windows
open System.Windows.Data
open System.Windows.Input
open System.ComponentModel
open System.Collections.ObjectModel
open RNExcel.Model
open RNExcel.Repository

type RNExcelHomeViewModel(accountRepository : AccountRepository)  =  
    inherit ViewModelBase()

    let mutable selectedAccount = 
        {Name=""; Role=""; Password=""; ExpenseLineItems = []}

    let mutable login                   = ""
    let mutable password                = ""
    let mutable document                = ""
    let mutable varCol                  = "A"
    let mutable valCol                  = "B"
    let mutable startRow                = "2"
    let mutable endRow                  = "0"
    let mutable convertButtonEnabled    = false
    let mutable loginExpander           = false

    let ApproveExpenseReport() = 
        let rc = new ExcWor()
        MessageBox.Show(sprintf "Excel said = %s" 
            (rc.make(document,varCol,valCol,startRow,endRow)) ) |> ignore

    new () = RNExcelHomeViewModel(new AccountRepository())

    member X.Accounts = 
        new ObservableCollection<AccountModel>(accountRepository.GetAll())

    member X.ApproveExpenseReportCommand = 
        new RelayCommand ((fun canExecute -> true), (fun action -> ApproveExpenseReport()))

    member X.FileOpen = 
        new RelayCommand ((fun canExecute -> true), (fun action -> 
            let dlg = new Microsoft.Win32.OpenFileDialog() 
            dlg.FileName    <- "Document"
            dlg.DefaultExt  <- ".xls"
            dlg.Filter      <- "Excel 2003 (.xls)|*.xls"
            dlg.FileOk.AddHandler (fun s e -> X.Document <- dlg.FileName)
            ignore <| dlg.ShowDialog() ))

    member X.LoginCommand =
        new RelayCommand((fun canExecute -> true),(fun action ->
                X.SelectedAccount <-
                    match
                        X.Accounts
                        |> Seq.filter (fun acc -> 
                            acc.Name        = login && 
                            acc.Password    = password) with
                        | s when Seq.isEmpty s -> 
                            X.ConvertButtonEnabled <- false
                            ignore <| MessageBox.Show(sprintf 
                                "User %s doesn't exist or password incorrect password" X.Login) 
                            {Name=""; Role=""; Password=""; ExpenseLineItems = []}
                        | s -> 
                            X.ConvertButtonEnabled <- true
                            X.LoginExpander <- false
                            Seq.head s

                X.Login     <- ""
                X.Password  <- "" ))

    member X.Login
        with get()      = login
        and set value   = 
            login <- value
            X.OnPropertyChanged "Login"

    member X.Password
        with get()      = password
        and set value   = 
            password <- value
            X.OnPropertyChanged "Password"

    member X.Document
        with get()      = document
        and set value   = 
            document <- value
            X.OnPropertyChanged "Document"

    member X.ConvertButtonEnabled
        with get()      = convertButtonEnabled
        and set v       = 
            convertButtonEnabled <- v
            X.OnPropertyChanged "ConvertButtonEnabled"

    member X.LoginExpander
        with get()      = loginExpander
        and set v       = 
            loginExpander <- v
            X.OnPropertyChanged "LoginExpander"

    member X.SelectedAccount 
        with get () = selectedAccount
        and set value = 
            selectedAccount <- value
            X.OnPropertyChanged "SelectedAccount"

    member X.VarCol 
        with get () = varCol
        and set value = 
            varCol <- value
            X.OnPropertyChanged "VarCol"

    member X.ValCol 
        with get () = valCol
        and set value = 
            valCol <- value
            X.OnPropertyChanged "ValCol"

    member X.StartRow 
        with get () = startRow
        and set value = 
            startRow <- value
            X.OnPropertyChanged "StartRow"

    member X.EndRow 
        with get () = endRow
        and set value = 
            endRow <- value
            X.OnPropertyChanged "EndRow"