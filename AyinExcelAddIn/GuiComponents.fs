namespace AyinExcelAddIn

open System
open System.Drawing
open System.Windows.Forms

module public Gui =
    type InputBoxResult<'a> =
        | Ok of 'a
        | Cancel

    let bondTickersCol = 
        let db = Db.nyabsDbCon.GetDataContext()
        let col = new AutoCompleteStringCollection()

        let symbols =
            query {
                for t in db.Intex.Tranches do
                    where (t.BbgTicker <> "")
                    select t.BbgTicker
                    distinct
            }
            |> Seq.toList

        symbols
        |> List.map (fun s -> col.Add(s.Trim()))
        |> ignore
        col
    //
    //
    //
    let public ErrorBox(msg : string) = MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore

    //
    //
    //
    let public SimpleInputBox (title : string) (label : string) (autoCompCol : AutoCompleteStringCollection option) =
        let res = ref (Ok "")
        let form =
            new Form(Text = title, FormBorderStyle = FormBorderStyle.FixedSingle, Height = 170, Width = 224,
                     BackColor = Color.White)
        let txtBox = new TextBox(Top = 57, Left = 32, Height = 23, Width = 160)

        if autoCompCol.IsSome then
            txtBox.AutoCompleteMode <- AutoCompleteMode.Suggest
            txtBox.AutoCompleteSource <- AutoCompleteSource.CustomSource
            txtBox.AutoCompleteCustomSource <- autoCompCol.Value
            txtBox.Text <- ""

        let lbl = new Label(Top = 27, Left = 32, Height = 23, Width = 160)
        lbl.Text <- label

        let okButton = new Button(Text = "Ok", Top = 98, Left = 127, Width = 65, BackColor = Color.LightGray)
        let cancelButton = new Button(Text = "Cancel", Top = 98, Left = 32, Width = 65, BackColor = Color.LightGray)

        okButton.Click.Add(fun _ ->
            res := Ok txtBox.Text
            form.Hide())

        cancelButton.Click.Add(fun _ ->
            res := Cancel
            form.Hide())

        // Create the form window
        form.Controls.Add(txtBox)
        form.Controls.Add(lbl)
        form.Controls.Add(okButton)
        form.Controls.Add(cancelButton)
        form.AcceptButton <- okButton
        form.ShowDialog() |> ignore
        !res

    //
    //
    //
    let public DateRangeInputBox(title : string) =
        let sd = new DateTime(2000, 1, 1)
        let ed = DateTime.Today
        let res = ref (Ok(sd, ed))
        let form =
            new Form(Text = title, FormBorderStyle = FormBorderStyle.FixedSingle, Height = 165, Width = 330,
                     BackColor = Color.White)

        let sdBox = new TextBox(Top = 27, Left = 150, Height = 23, Width = 160)
        let edBox = new TextBox(Top = 55, Left = 150, Height = 23, Width = 160)

        sdBox.Text <- sd.ToString("MM/dd/yyyy")
        edBox.Text <- ed.ToString("MM/dd/yyyy")

        let sdLabel = new Label(Top = 27, Left = 10, Height = 23, Width = 160)
        let edLabel = new Label(Top = 55, Left = 10, Height = 23, Width = 160)

        sdLabel.Text <- "Start date"
        edLabel.Text <- "End date"

        let okButton = new Button(Text = "Ok", Top = 89, Left = 180, Width = 65, BackColor = Color.LightGray)
        let cancelButton = new Button(Text = "Cancel", Top = 89, Left = 85, Width = 65, BackColor = Color.LightGray)

        okButton.Click.Add(fun _ ->
            try
                let sdInput = DateTime.Parse(sdBox.Text)
                let edInput = DateTime.Parse(edBox.Text)
                if sdInput > edInput then ErrorBox "Start date must be before end date!"
                res := Ok(sdInput, edInput)
                form.Hide()
            with :? FormatException -> ErrorBox "Invalid date")

        cancelButton.Click.Add(fun _ ->
            res := Cancel
            form.Hide())

        // Create the form window
        form.Controls.Add(sdBox)
        form.Controls.Add(edBox)
        form.Controls.Add(sdLabel)
        form.Controls.Add(edLabel)
        form.Controls.Add(okButton)
        form.Controls.Add(cancelButton)
        form.AcceptButton <- okButton
        form.ShowDialog() |> ignore
        !res

    let public DatePickerInputBox p func =
        let form = new Form(FormBorderStyle = FormBorderStyle.None, Height = 206, Width = 180, BackColor = Color.White)
        let calendar = new MonthCalendar(Top = 0, Left = 0, Height = 180, Width = 180)

        calendar.Margin <- new Padding(0)
        calendar.MinDate <- new DateTime(2000, 1, 1)
        calendar.MaxDate <- DateTime.Today
        calendar.ShowToday <- false
        calendar.ShowTodayCircle <- false
        calendar.DateSelected.Add(fun _ ->
            form.Hide()
            func calendar.SelectionStart)

        let cancelButton =
            new Button(Text = "Cancel", Top = 180, Left = 0, Height = 25, Width = 180, BackColor = Color.LightGray)

        cancelButton.Click.Add(fun _ -> form.Hide())
        form.Controls.Add(calendar)
        form.Controls.Add(cancelButton)
        form.Show()
        form.Location <- p

