//---------------------------------------------------------------------------------------
// This class calculates the flag InvalidThrPos_B.
// The flag indicates an unallowable strongly pressed throttle pedal and leads to
// deactivation of HDC. Purpose: misuse prevention. 
//---------------------------------------------------------------------------------------


// calculate Throttle decrement according to parameter C_tThrPosInvalid4FullThr
// avoid division by zero
if (Enable_B)
{
  // limit negative throttle gradient towards actual ThrottlePosition
  ThrPosInvldFlt = (ThrPosInvldFlt - C_HdcThrDecrPerCycle).max(ThrottlePosition);
  if ((AxRoadSlopeF < -C_HdcActSlopeThr)	//Downhill driving
	  ||((AxRoadSlopeF > C_HdcActSlopeThr)
	      &&(  ((GearDirInfo != VehMovDir_Forward)&&(WssDirInfo == VehicleDirectionInfo_Forward))		 
			 ||((GearDirInfo != VehMovDir_Backward)&&(WssDirInfo == VehicleDirectionInfo_Backward)))	//Rollback with uphill gear or N gear
		  )
	  )
  {
	  tDownhillDetected ++;
  }
  else if (AxRoadSlopeF > -(C_HdcActSlopeThr - 0.2))  //reset tDownhillDetected when not downhill or not rollback with uphill gear or N gear
  {
	  tDownhillDetected = 0;
  }
  
  if (tDownhillDetected > C_HdcActSlopeTimeDetThr)
  {
	  DownhillDetected = true;
  }
  else
  {
	  DownhillDetected = false;
  }
  
  if (DownhillDetected)
  {
	   // if filtered ThrottlePosition is over the allowed threshold
	   if (ThrPosInvldFlt > C_HdcThrottlePosMax)
       {
           InvalidThrPos_B = true;
       }
       else
       {
           InvalidThrPos_B = false;
       }
  }
  else
  {
	  InvalidThrPos_B = true;
  }
}

else
{
  // reset InvalidThrPos_B
  ThrPosInvldFlt = 0.0;
  InvalidThrPos_B = false;
}
