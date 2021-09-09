// collect last multiple vehicle data
If(finishMultipleVehicles=true, // add the last vehicle and add it to the collection
// check format of reg before collecting the data -------------------------------
If(!IsBlank(inpReg),
UpdateContext( // set variables
    {
        varKeyedReg:Upper(inpReg.Text),
        varKeyedRegLength:Len(inpReg.Text),
        varKeyedRegSpacePosition:Find(" ",inpReg.Text,1),
        varKeyedRegNew:""
    }
);
// if seven characters - no space - then update reg
If(varKeyedRegLength=7,
    UpdateContext({varKeyedRegNew:Replace(varKeyedReg,1,4,Mid(varKeyedReg,1,4) & " ")}),
    UpdateContext({varKeyedRegNew:varKeyedReg}) // else
    )
); // end if format reg ----------------------------------------------------------------
// This is not a change
// collect multiple vehicle data
Collect(
        colMultipleVehicles,
    {
        Make: inpMake.Text,
        Model: inpModel.Selected.Result,
        Colour: dropColour.Selected.Value,
        VantageRef:Value(inpVantageRef.Text),
        Chassis:Right(inpFullChassis.Text,5),
        FullChassis:inpFullChassis.Text,
        RegNumber:varKeyedRegNew,
        WestWayPo:inpWestWayPo.Text,
        BodyType:dropBodyType.Selected.Value,
        WheelBase:dropWheelBase.Selected.Value,
        RoofHeight:dropRoofHeight.Selected.Value,
        FittingAddress:inpFittingAddress.Text,
        SideLoadingDoors:radioSideLoadingDoor.Selected.Value,
        RearDoors:radioRearDoor.Selected.Value,
        Glass:radioGlass.Selected.Value,
        NissanOrderNumber:inpNissanOrderNumber.Text,
        DueDate:dateDueDate.SelectedDate
    }
);
// patch multiple vehicle data
If(finishMultipleVehicles=true,
    ClearCollect(colFinalMultipleVehicles,RenameColumns(colMultipleVehicles,"Make","Title"));
    Collect('West Way Vehicle List',colFinalMultipleVehicles);
/*
// send email to the group
    Office365Outlook.SendEmailV2("westway@vantagevcltd.co.uk",
    "New vehicles on the West Way Portal",
     "The following vehicles have been added to the West Way portal" &
     "<br><a href='https://apps.powerapps.com/play/410f7538-fb74-48eb-932c-118fe877492e?tenantId=5ff4a4be-a67a-4b4c-a948-ad6894073dc5'>West Way Portal</a>" &
            Concat(colFinalMultipleVehicles,            
                            "<br><br>Model: " & Model &
                            "<br>Reg: " & RegNumber &
                            "<br>Chassis: " & Chassis),
                            {From:User().FullName}
    )
    )
);
*/
// create message details
UpdateContext({collateMessageDetails:"The following vehicles have been added to the West Way portal:" & 
                            Concat(colFinalMultipleVehicles,
                            "<br><br>Model: " & Model &
                            "<br>Reg: " & RegNumber &
                            "<br>Chassis: " & Chassis &
                            "<br>Full Chassis: " & FullChassis                            
							) &
                            "<br>Spec: " & Concat(colSpecItems,SpecItem & " x " & SpecQuantity & "<br>") &
							"<br><br><a href='https://apps.powerapps.com/play/2e197de3-3568-40b2-bac2-3ef0dffa98fa?tenantId=5ff4a4be-a67a-4b4c-a948-ad6894073dc5'>West Way Portal</a>"
				});
                
// patch email list with fitting date data
Patch('West Way Emails',Defaults('West Way Emails'),
    {
    Message:collateMessageDetails,
    Title:varUserName,
    MessageType:"New Vehicles Added to Portal"
    }
)
)
);
// -----------------------------------------------------------------------------------------------------------------

If(finishMultipleVehicles=false, // only one vehicle added
// check format of reg before collecting the data -------------------------------
If(!IsBlank(inpReg),
UpdateContext( // set variables
    {
        varKeyedReg:Upper(inpReg.Text),
        varKeyedRegLength:Len(inpReg.Text),
        varKeyedRegSpacePosition:Find(" ",inpReg.Text,1),
        varKeyedRegNew:""
    }
);
// if seven characters - no space - then update reg
If(varKeyedRegLength=7,
    UpdateContext({varKeyedRegNew:Replace(varKeyedReg,1,4,Mid(varKeyedReg,1,4) & " ")}),
    UpdateContext({varKeyedRegNew:varKeyedReg}) // else
    )
); // end if format reg ----------------------------------------------------------------

// patch single vehicle data
Patch(
    'West Way Vehicle List',
    Defaults('West Way Vehicle List'),
    {
        Make: inpMake.Text,
        Model: inpModel.Selected.Result,
        Colour: dropColour.Selected.Value,
        VantageRef:Value(inpVantageRef.Text),
        Chassis:Right(inpFullChassis.Text,5),
        FullChassis:inpFullChassis.Text,
        RegNumber:varKeyedRegNew,
        WestWayPo:inpWestWayPo.Text,
        BodyType:dropBodyType.Selected.Value,
        WheelBase:dropWheelBase.Selected.Value,
        RoofHeight:dropRoofHeight.Selected.Value,
        FittingAddress:inpFittingAddress.Text,
        SideLoadingDoors:radioSideLoadingDoor.Selected.Value,
        RearDoors:radioRearDoor.Selected.Value,
        Glass:radioGlass.Selected.Value,
        NissanOrderNumber:inpNissanOrderNumber.Text,
        DueDate:dateDueDate.SelectedDate
    }
);
/*
// send email for single vehicle
Office365Outlook.SendEmailV2("westway@vantagevcltd.co.uk",
    "New vehicle on the West Way Portal",
     "The following vehicle has been added to the West Way portal" &
                            "<br><br>Model: " & inpModel.Selected.Result & 
                            "<br>Chassis: " & Right(inpFullChassis.Text,6) &
                            "<br>Reg Number: " & inpReg.Text &
                            "<br>Body type: " & dropBodyType.Selected.Value &
                            "<br>Wheelbase: " & dropWheelBase.Selected.Value &
                            "<br>Roof Height: " & dropRoofHeight.Selected.Value &
                            "<br>Side Loading Doors: " & radioSideLoadingDoor.Selected.Value &
                            "<br>Glass: " & radioGlass.Selected.Value &
                            "<br>Rear Doors: " & radioRearDoor.Selected.Value &
                            "<br>Fitting Address: " & inpFittingAddress.Text &
                            "<br><br><a href='https://apps.powerapps.com/play/410f7538-fb74-48eb-932c-118fe877492e?tenantId=5ff4a4be-a67a-4b4c-a948-ad6894073dc5'>West Way Portal</a>",
                            {From:User().FullName,Cc:User().Email}
    
)
);
*/
// patch query message data
UpdateContext({collateMessageDetails:"The following vehicle has been added to the West Way portal:" &
                            "<br><br>Model: " & inpModel.Selected.Result & 
                            "<br>Chassis: " & Right(inpFullChassis.Text,6) &
                            "<br>Full Chassis: " & inpFullChassis.Text &
                            "<br>Reg Number: " & inpReg.Text &
                            "<br>Body type: " & dropBodyType.Selected.Value &
                            "<br>Wheelbase: " & dropWheelBase.Selected.Value &
                            "<br>Roof Height: " & dropRoofHeight.Selected.Value &
                            "<br>Side Loading Doors: " & radioSideLoadingDoor.Selected.Value &
                            "<br>Glass: " & radioGlass.Selected.Value &
                            "<br>Rear Doors: " & radioRearDoor.Selected.Value &
                            "<br>Fitting Address: " & inpFittingAddress.Text &
                            "<br>PO Number: " & inpWestWayPo.Text &
                            "<br>Spec: " & Concat(colSpecItems,SpecItem & " x " & SpecQuantity & "<br>") &                           
                            "<br><br><a href='https://apps.powerapps.com/play/2e197de3-3568-40b2-bac2-3ef0dffa98fa?tenantId=5ff4a4be-a67a-4b4c-a948-ad6894073dc5'>West Way Portal</a>"
                            });

Patch('West Way Emails',Defaults('West Way Emails'),
    {
    Message:collateMessageDetails,
    Title:varUserName,
    MessageType:"New vehicle on the portal"
    }
)
);
// patch spec data
Collect('West Way Specs',colSpecItems);

// reset
Reset(inpMake);Reset(inpModel);Reset(dropColour);Reset(inpFullChassis);Reset(inpReg);UpdateContext({varKeyedRegNew:Blank()});Reset(inpWestWayPo);Reset(inpNissanOrderNumber);
Reset(dropBodyType);Reset(dropWheelBase);Reset(dropRoofHeight);Reset(inpFittingAddress);
Reset(radioGlass);Reset(radioRearDoor);Reset(radioSideLoadingDoor);Reset(dropNumberOfVehicles);Reset(dateDueDate);
Reset(radioPlylining);UpdateContext({varPlylining:false,varPlyliningPosition:Blank()});
Reset(radioSupplyTowbar);UpdateContext({varTowbar:false,varTowbarSupply:Blank()});
Reset(radioTowbarElectricsSupply);Reset(radioTowbarElectricsType);UpdateContext({varTowbarElectrics:false,varTowbarElectricsType:Blank(),varTowbarElectricSupply:Blank()});
Reset(radioSupplyBedliner);UpdateContext({varBedliner:false,varBedlinerSupply:Blank()});
Reset(radioSupplyMats);UpdateContext({varMats:false,varMatsSupply:Blank()});
Reset(radioSupplyTailgateAssist);UpdateContext({varTailgateAssist:false,varTailgateAssistSupply:Blank()});
Reset(dropTruckmanSize);Reset(radioTruckmanSupply);UpdateContext({varTruckman:false,varTruckmanSize:Blank()});
Clear(colSpecItems);Clear(colMultipleVehicles);Clear(colFinalMultipleVehicles);Navigate(Overview)// collect last multiple vehicle data
