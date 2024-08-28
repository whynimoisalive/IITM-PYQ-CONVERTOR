param (
    [string]$inputPdf,
    [string]$outputPdf
)

# Ensure the correct printer is set
$printerName = "Microsoft Print to PDF"

# Set up the printing process
$PrintJob = New-Object -ComObject WScript.Network
$PrintJob.SetDefaultPrinter($printerName)

# Use COM to create a WScript.Shell object
$shell = New-Object -ComObject WScript.Shell

# Open the PDF in the default viewer (e.g., Adobe Acrobat Reader)
$shell.Run('"' + $inputPdf + '"')

Start-Sleep -Seconds 1 # Wait for the PDF to open

# Simulate keystrokes to print the document
$shell.SendKeys("^p") # Ctrl + P to open the print dialog
Start-Sleep -Seconds 1.5 # Wait for the dialog to open
$shell.SendKeys("{TAB 6}") # Tab to color options (adjust number of tabs as needed)
$shell.SendKeys("{DOWN}") # Switch to "Black & White"

$shell.SendKeys("{ENTER}") # Confirm settings
Start-Sleep -Seconds 0.5
$shell.SendKeys("{TAB 3}")
