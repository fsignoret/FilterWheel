// --------------------------------------------------------------------------------
//
// Arduino code for ArduinoFilterWheel
//
// Description:	Arduino code for a filter wheel assembled by Laurent Brunetto (GAPRA),
//              based on Arduino
//
// Protocol :   Setup phase moves the filter wheel to the home position (first filter)
//              With command GETFILTER#, the client gets the current filter position (beginning at 0, followed by #)
//              With command SETFILTER[0-9]#, the client sets the current filter position
//                The command returns the new filter position, followed by #.
//
// Author:		Florian Signoret <flo@signoret.org>
//
// Edit Log:
//
// Date			Who	                Vers	Description
// -----------	---	                -----	---------------------------------------
// 26-Aug-2014	Florian Signoret	  1.0.0	  Initial edit
// 29-Dec-2015  Florian Signoret    1.0.1   Adapted to DC Motor
// 23-Feb-2016  Florian Signoret    1.0.2   Send current position after initialization
// --------------------------------------------------------------------------------
//

#include <AFMotor.h>

const String GETFILTER_COMMAND = "GETFILTER";
const String SETFILTER_COMMAND = "SETFILTER";
const char COMMAND_DELIMITER_CHAR = '#';
const long SERIAL_PORT_SPEED = 9600; // 9600 bps
const int NUMBER_OF_POSITIONS = 5;
const int MOTOR_SPEED = 48;
const int MAX_SEARCH_STEPS = 300;
const int INITIAL_DELAY = 300;
const int NEXT_MOVE_DELAY = 50;
const int OPTICAL_CAPTOR_PIN = 16;
const int HALL_CAPTOR_PIN = 14;
const int HALL_CAPTOR_POSITION = 0; // Hall captor is on Filter #0 position

AF_DCMotor motor(1, MOTOR12_2KHZ);
int _currentPosition = HALL_CAPTOR_POSITION; // At startup, we want the home position
bool _isLost = true; // true if the last captor search failed, or position is unknown (at startup)

void setup()
{
  // Initialize serial port
  Serial.begin(SERIAL_PORT_SPEED);
  Serial.flush();

  // Initialize motor and captors
  motor.setSpeed(MOTOR_SPEED);
  pinMode(OPTICAL_CAPTOR_PIN, INPUT);
  pinMode(HALL_CAPTOR_PIN, INPUT);

  //MoveFilterWheel(_currentPosition);
  
  SendCurrentPosition();
}

void loop()
{
  String command;

  if (Serial.available() > 0)
  {
    // Read the received command on Serial
    command = Serial.readStringUntil(COMMAND_DELIMITER_CHAR);

    // Process according to command
    if (command == GETFILTER_COMMAND)
    {
      SendCurrentPosition();
    }
    else if (command.startsWith(SETFILTER_COMMAND))
    {
      int targetPosition = command.substring(SETFILTER_COMMAND.length()).toInt();
      // Call twice in order to be accurately positioned
      MoveFilterWheel(targetPosition);
      MoveFilterWheel(targetPosition);
      
      SendCurrentPosition();
    }
  }
}

void SendCurrentPosition()
{
  Serial.print(_currentPosition);
  Serial.println(COMMAND_DELIMITER_CHAR);
}

bool MoveFilterWheel(int targetPosition)
{
  /*Serial.print("MoveFilterWheel(");
  Serial.print(targetPosition);
  Serial.println(")");*/
  
  // Check that the targetPosition is valid
  if (targetPosition < 0 || targetPosition > NUMBER_OF_POSITIONS - 1)
    return false;

  if (_isLost)
  {
    //Serial.println("Lost!");
    bool homeFound = FindHomeCaptor();
    if (!homeFound)
      return false;
  }

  // Calculate the number of position changes that needs to be performed to reach the target position
  int numberOfPositionChanges = targetPosition - _currentPosition;
  int motorDirection = FORWARD;
  if (numberOfPositionChanges < 0)
    numberOfPositionChanges += NUMBER_OF_POSITIONS;
    
  // Optimize the move: if the target position is on the other half, go backward
  if (2 * numberOfPositionChanges > NUMBER_OF_POSITIONS)
  {
    numberOfPositionChanges = NUMBER_OF_POSITIONS - numberOfPositionChanges;
    motorDirection = BACKWARD;
  }

  // Perform the moves
  for (int i = 0; i < numberOfPositionChanges; ++i)
  {
    int nextPosition = motorDirection == FORWARD ? _currentPosition + 1 : _currentPosition - 1;
    if (nextPosition < 0)
      nextPosition += NUMBER_OF_POSITIONS;
    bool captorFound = FindNextCaptor(motorDirection, nextPosition % NUMBER_OF_POSITIONS);
    if (!captorFound)
      return false;
  }
  
  // If no position changes needed, make sure we are accurately positionned
  if (numberOfPositionChanges == 0)
    return FindNextCaptor(FORWARD, targetPosition);

  return true;
}

bool FindHomeCaptor()
{
  //Serial.println("FindHomeCaptor()");
  return FindNextCaptor(FORWARD, HALL_CAPTOR_POSITION);
}

/*
 * Moves the wheel forward in order to reach the next captor
 * Updates the new position attribute if the captor was found.
 * Returns true if the captor was found, false otherwise
 */
bool FindNextCaptor(int motorDirection, int newPosition)
{
  int captorPin = GetCaptorPinFromPosition(newPosition);
  
  /*Serial.print("FindNextCaptor(");
  Serial.print(motorDirection);
  Serial.print(", ");
  Serial.print(newPosition);
  Serial.println(")");*/
  
  // Try to leave the current position
  // Backward if we already are on the target position
  // Forward otherwise
  if (newPosition == _currentPosition)
    motor.run(BACKWARD);
  else
    motor.run(motorDirection);
  
  // Mininum delay
  delay(INITIAL_DELAY);
  
  // Leave the captor position
  while (IsCaptorDetected(captorPin))
    delay(NEXT_MOVE_DELAY);
  motor.run(RELEASE);

  // Then reach the next position
  motor.run(motorDirection);
  
  boolean isPositionned = false;
  int numberOfSearchSteps = 1;
  while (!isPositionned && numberOfSearchSteps < MAX_SEARCH_STEPS)
  {
    delay(NEXT_MOVE_DELAY);
    isPositionned = IsCaptorDetected(captorPin);
    ++numberOfSearchSteps;
  }
  
  // Stop the motor
  motor.run(RELEASE);
  
  if (isPositionned)
    _currentPosition = newPosition;

  _isLost = !isPositionned;

  return isPositionned;
}

int GetCaptorPinFromPosition(int pos)
{
  if (pos == HALL_CAPTOR_POSITION)
    return HALL_CAPTOR_PIN;
  return OPTICAL_CAPTOR_PIN;
}

bool IsCurrentCaptorDetected()
{
  return IsCaptorDetected(GetCaptorPinFromPosition(_currentPosition));
}

bool IsCaptorDetected(int captorPin)
{
  return digitalRead(captorPin) == LOW;
}

bool IsOpticalCaptorDetected()
{
  return IsCaptorDetected(OPTICAL_CAPTOR_PIN);
}

bool IsHallCaptorDetected()
{
  return IsCaptorDetected(HALL_CAPTOR_PIN);
}
