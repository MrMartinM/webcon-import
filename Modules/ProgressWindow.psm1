# ProgressWindow.psm1
# Module for displaying progress in a Windows Forms window

function Show-ProgressWindow {
    <#
    .SYNOPSIS
    Creates and displays a progress window with a progress bar
    
    .PARAMETER TotalRows
    Total number of rows to process
    
    .EXAMPLE
    $progressWindow = Show-ProgressWindow -TotalRows 100
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [int]$TotalRows
    )
    
    # Load Windows Forms assembly
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    # Create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Webcon Import Progress"
    $form.Size = New-Object System.Drawing.Size(700, 220)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true
    $form.ShowInTaskbar = $true
    
    # Create progress bar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 20)
    $progressBar.Size = New-Object System.Drawing.Size(650, 30)
    $progressBar.Minimum = 0
    # Ensure Maximum is at least 1 to avoid division issues
    $progressBar.Maximum = [Math]::Max($TotalRows, 1)
    $progressBar.Style = "Continuous"
    $progressBar.Value = 0
    $form.Controls.Add($progressBar)
    
    # Create status label (use TextBox for multiline text support with word wrapping)
    $statusLabel = New-Object System.Windows.Forms.TextBox
    $statusLabel.Location = New-Object System.Drawing.Point(20, 60)
    $statusLabel.Size = New-Object System.Drawing.Size(650, 35)
    $statusLabel.Multiline = $true
    $statusLabel.ReadOnly = $true
    $statusLabel.BorderStyle = "None"
    $statusLabel.BackColor = $form.BackColor
    $statusLabel.WordWrap = $true
    $statusLabel.TabStop = $false
    $statusLabel.Text = "Ready to process $TotalRows rows..."
    $statusLabel.ScrollBars = "None"
    $form.Controls.Add($statusLabel)
    
    # Create details label
    $detailsLabel = New-Object System.Windows.Forms.Label
    $detailsLabel.Location = New-Object System.Drawing.Point(20, 105)
    $detailsLabel.Size = New-Object System.Drawing.Size(650, 20)
    $detailsLabel.Text = "Waiting to start processing..."
    $form.Controls.Add($detailsLabel)
    
    # Create close button (initially disabled)
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(300, 135)
    $closeButton.Size = New-Object System.Drawing.Size(100, 30)
    $closeButton.Text = "Close"
    $closeButton.Enabled = $false
    $closeButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($closeButton)
    
    # Show the form (non-blocking)
    # Note: Removed Add_Shown event handler to avoid null reference exception
    $form.Show()
    $form.Activate()
    $form.Refresh()
    
    # Return an object with methods to update the progress
    $script:progressWindowObj = @{
        Form = $form
        ProgressBar = $progressBar
        StatusLabel = $statusLabel
        DetailsLabel = $detailsLabel
        CloseButton = $closeButton
        TotalRows = $TotalRows
    }
    
    # Add update method
    $script:progressWindowObj | Add-Member -MemberType ScriptMethod -Name UpdateProgress -Value {
        param(
            [int]$Processed,
            [string]$CurrentRow,
            [int]$SuccessCount,
            [int]$ErrorCount,
            [int]$SkippedCount
        )
        
        # Update progress bar (ensure we don't exceed maximum)
        $maxValue = [Math]::Max($this.TotalRows, 1)
        $this.ProgressBar.Value = [Math]::Min($Processed, $maxValue)
        
        # Calculate percentage
        $percent = if ($this.TotalRows -gt 0) {
            [Math]::Round(($Processed / $this.TotalRows) * 100, 1)
        } else { 
            if ($Processed -gt 0) { 100 } else { 0 }
        }
        
        # Update status label (fix string interpolation for TotalRows)
        $totalRowsValue = $this.TotalRows
        $this.StatusLabel.Text = "Processing: $CurrentRow | Progress: $Processed / $totalRowsValue ($percent%)"
        
        # Update details label with formatted counts
        $this.DetailsLabel.Text = "Successful: $SuccessCount | Errors: $ErrorCount | Skipped: $SkippedCount"
        
        # Refresh the form to update UI
        try {
            $this.Form.Refresh()
            [System.Windows.Forms.Application]::DoEvents()
        } catch {
            # Silently ignore refresh errors (form might be disposed)
        }
    }
    
    # Add close method
    $script:progressWindowObj | Add-Member -MemberType ScriptMethod -Name Close -Value {
        $this.CloseButton.Enabled = $true
        $this.CloseButton.Text = "Close"
        $this.StatusLabel.Text = "Processing Complete!"
        $this.Form.Refresh()
    }
    
    return $script:progressWindowObj
}

Export-ModuleMember -Function Show-ProgressWindow

