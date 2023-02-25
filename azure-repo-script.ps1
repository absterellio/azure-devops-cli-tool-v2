Add-Type -assembly System.Windows.Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Repo Report'
$form.Size = New-Object System.Drawing.Size(600,400)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(200,200)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(290,200)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

# personal access token
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(150,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Personal access token:'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(150,40)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox)

# repo files to query
$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(150,70)
$label1.Size = New-Object System.Drawing.Size(280,40)
$label1.Text = 'Enter a text file to read from. (Note: Specify each repo you would like to query on a new line of the file):'
$form.Controls.Add($label1)

$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = New-Object System.Drawing.Point(150,110)
$textBox1.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox1)

# specify output file
$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(150,140)
$label2.Size = New-Object System.Drawing.Size(280,20)
$label2.Text = 'Please specify a path for the output csv file:'
$form.Controls.Add($label2)

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(150,160)
$textBox2.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox2)

$form.Topmost = $true

$form.Add_Shown({$textBox.Select()})
$form.Add_Shown({$textBox1.Select()})
$form.Add_Shown({$textBox2.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $pat = $textBox.Text
    $fileName = $textBox1.Text
    $location = $textBox2.Text
    $projectName = "ex-projectName"

    # authenticate based on the PAT
    echo $pat | az devops login --organization ex-organization-url

    # set csv file headers
    Set-Content -Path $location -Value '"PR Number","Date Closed","Repo Name","Created By","Reviewers","Work Item Id"'
    $repoNames = Get-Content $fileName
    # for each repo, get pull request data
    foreach ($repo in $repoNames) {

        # grab the completed PR's that went into main for this repo
        $prListObj = az repos pr list --project $projectName --repository $repo --status completed --target-branch main --include-links | ConvertFrom-Json 

        foreach ($pr in $prListObj) {
            $prNumber = $pr.pullRequestId
            $workItem = (az repos pr work-item list --id $pr.pullRequestId) | ConvertFrom-Json
            $workItemId = $workItem.id
            $createdBy = $pr.createdBy.displayName
            $reviewedBy = $pr.reviewers.displayName
            $date = $pr.closedDate

            # get the reviewers display string
            if($reviewedBy[1]) {
                $reviewedByString = "$($reviewedBy[1]) on behalf of $($reviewedBy[0])"
            } else {
                $reviewedByString = $reviewedBy[0]
            }

            # output each pr in the csv file
            Add-Content -Path $location -Value "$prNumber,$date,$repo,$createdBy,$reviewedByString,$workItemId"
        }
    }

}



