# writerbird-installer

Within the distribution folder, the three files specific to the installer are:
- writerbird-installer.py: the python script that runs the installer
- writerbird-word-reg.txt: Microsoft provided registry script to direct MS Word towards add-in folder locations. See the [Microsoft Site]( https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) for more information: 
- writerbird-manifest.xml: the add-in specific information read by Word

The remainder of the files are generic files supplied from [Python](https://www.python.org/downloads/windows/) that allow the installer script to be run on the user computer. 
