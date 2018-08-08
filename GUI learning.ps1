#Load Assembly and Library
Add-Type -AssemblyName PresentationFramework

#Simple PopUp box, first String is text, second String is title
#[System.Windows.MessageBox]::Show("Cannot retrive event logs from server","Lame")

#XAML form designed using Vistual Studio
[xml]$Form = @"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Omega Simplified" Height="450" Width="800">
    <Grid>
        <Button Name="Button_Bulk_Permissions" Content="Mailbox Permissions - Bulk" HorizontalAlignment="Left" Margin="50,100,0,0" VerticalAlignment="Top" Width="166"/>
        <Button Content="Server CPU Utilization" HorizontalAlignment="Left" Margin="50,50,0,0" VerticalAlignment="Top" Width="166"/>
        <Button Content="Data Sharing Service" HorizontalAlignment="Left" Margin="50,150,0,0" VerticalAlignment="Top" Width="166"/>
        <Button Content="Out of Office" HorizontalAlignment="Left" Margin="50,200,0,0" VerticalAlignment="Top" Width="166"/>

    </Grid>
</Window>
"@

#Create a form
$XMLReader = (New-Object System.Xml.XmlNodeReader $Form)
$XMLForm = [Windows.Markup.XamlReader]::Load($XMLReader)

$Event_Bulk_Permissions = $XMLForm.FindName('Button_Bulk_Permissions')

$Event_Bulk_Permissions.add_click({
    Write-Host("You have clicked Mailbox Permissions - Bulk button")
})


#Show XMLform
$XMLForm.ShowDialog()