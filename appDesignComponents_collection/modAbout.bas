Attribute VB_Name = "modAbout"
'          _________
'_________| about    \___ changes ____________________________________________________________________________________________________________________________
'
'
'        ____ [ collection ] _________________________________________________________________________________________________________________________________
'       |
'       |   > name: appDesingComponents Collection
'       |   > version: 1.4
'       |   > price: free for any kind of use
'       |______________________________________________________________________________________________________________________________________________________
'
'
'
'        ____ [ authors ] _____________________________________________________________________________________________________________________________________
'       |
'       |   > ivan stimac, croatia
'       |       mail: ivan.stimac@po.htnet.hr, flashboy01@gmail.com
'       |   >
'       |       mail:
'       |   >
'       |       mail:
'       |______________________________________________________________________________________________________________________________________________________
'
'
'
'        ____ [ thanks ] ______________________________________________________________________________________________________________________________________
'       |
'       |   > Ariad Software - ascPaintEffects class, modFile
'       |   > Mark Gordon - power resize
'       |   >
'       |_______________________________________________________________________________________________________________________________________________________
'
'
'
'        ____ [ please ] _______________________________________________________________________________________________________________________________________
'       |
'       |   > rate it
'       |   > report bugs
'       |   > add your components there and share them with us
'       |     '- if you do that, add also your name as author
'       |     '- or send your components to me and i will add them
'       |           and of course add your name as author
'       |_______________________________________________________________________________________________________________________________________________________
'
'
'
'
'
'
'
'                    ___________
'_________ about ___| changes    \______________________________________________________________________________________________________________________________
'
'        ____ [ 07.20.2004. ] _________________________________________________________________________________________________________________________________
'       |
'       |  <> Thanks Light Templer for reported errors:
'       |     ----------------------------------------------
'       |     "Comment From: Light Templer
'       |        Comment: Got error on opening demo form (Starts with property 'ForeColor',
'       |        seems none of the values are saved ...)
'       |        When I use an empty form, putting the controls on it all I got are empty rectangles ...
'       |      Suggestion: Always use default values on  PropertyRead.
'       |        Any more tipps?
'       |        Regards -LiTe"
'       |  <> NEW
'       |     ----------------------------------------------
'       |   > add 3 new styles to Tab control, and property page for it
'       |   > you can now edit tab control at desing time and see full results
'       |   > there is no more only empty rectangles, you have now sample button on controls ItemList and ImageMenu
'       |______________________________________________________________________________________________________________________________________________________
'
'
'
'
'        ____ [ 07.19.2004. ] _________________________________________________________________________________________________________________________________
'       |
'       |   > On imageMenu sub UserControl_MouseMove : "fix redraw whenever mouse move what cause slow
'       |                                               hover effect"
'       |   > Add new evetns to Tab, ItemList and ImageMenu controls (mouse move/down/up, OLEDragDrop1...)
'       |______________________________________________________________________________________________________________________________________________________
'
'
'
'        ____ [ 07.18.2004. ] _________________________________________________________________________________________________________________________________
'       |
'       |  <> Thanks Roger Gilchrist for reported errors:
'       |     ----------------------------------------------
'       |   > On ItemList control in Sub imgIc_MouseMove : If Index <> hoverIND And isInList(lstDisabled, i) <> True  "(i must be Index)"
'       |   > On frame control in ReadProperties : "BCControl should be BC"
'       |   > On ImageMenu control in Sub UserControl_WriteProperties : "FitmCDisabled should be itmFCDisabled"
'       |   > On ImageList control in Sub RemoveItem then line :  lstImages.Remove itmIndex + 1 makes an invalid cal to lstImages a control that doesn't exist
'       |                                                           "code line must be deleted"
'       |______________________________________________________________________________________________________________________________________________________
'
'
'
'        ____ [ 07.16.2006. ] _________________________________________________________________________________________________________________________________
'       |
'       |   > fix imagemenu and itemlist controls:
'       |     '- now can draw icon type images
'       |______________________________________________________________________________________________________________________________________________________
'
'
'
'        ____ [ 07.15.2006. ] _________________________________________________________________________________________________________________________________
'       |
'       |   > rename imageList to imageMenu
'       |   > create imageList component
'       |   > fix itemList and imageMenu component:
'       |     '- image drawing
'       |______________________________________________________________________________________________________________________________________________________
'
'
'
'
'
'
