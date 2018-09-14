# ExcelDna-MVVM
Databinding enabled MVVM Adapter for Excel-DNA

## Bind To your ViewModel
Declare WPF like Databinding in your .dna - File.
```xml
...
<group id='seTest' label='Test'>
  <button id='testButton' label='My cool databound Button' onAction='{Binding TestCommand}' 
          getEnabled='{Binding ButtonEnabled}' getVisible='{Binding ButtonVisible}'/>  
</group>
...
```

Declare Viewmodels (Declaring Assembly have to reside in the same App-Domain, the ExcelDna-MVVM Assembly is loaded in), like this

```cs
class TestAppVM : IAppVM
{
    public ICommand TestCommand { get; set; }
    public bool ButtonEnabled { get; set; }
    public bool ButtonVisible { get; set; }
  
  public TestAppVM()
  {
    TestCommand  = new RelayCommand((id) =>
      {
          MessageBox.Show($"Hello from Button '{id}'");//says: Hello from Button 'testButton' 
      }, (o) => { return true; });
  }
}
```
ExcelDna-MVVM searches for implementations of the following interfaces:
* IAppVM (Created per Application)
* IWorkbookVM (Created per Workbook)
* ISheetVM (Created per Worksheet)

If found, ExcelDna-MVVM creates them and creates Databinding for them.

So this Library enriches Excel-DNA by Databinding Possibilities, known from WPF.
This enables Addins, using the widespread MVVM-Pattern, most GUI-Developer are familiar with. 
