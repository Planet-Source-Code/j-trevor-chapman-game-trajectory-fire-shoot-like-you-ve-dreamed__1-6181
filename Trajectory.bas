Attribute VB_Name = "Trajectory"
'   A TRAJECTORY MODULE BY: J TREVOR CHAPMAN. PLEASE CREDIT ME IF USED.
'       AND I WOULD LOVE TO SEE THE PROJECTS THAT YOU USE THIS FOR
'           MODULE FOR. VOTE FOR ME! Feb 20 2000, 16 yrs old
'               THANKS AND ENJOY.

Option Explicit
'------------------------------------------------------------------
Public Const Pi As Double = 3.14159265358979        'Needed Math
Public Const Radians As Double = (2 * Pi) / 360     'Needed Math
'-------------------------------------------------------------------
    'The following are Vital vars. for the calculations.
    'They are kept public in case you need to access them.

Public A As Single                     'Angle
Public T As Single                     'Timer
Public X As Integer                    'The Y axis increment properties
Public Y As Integer                    'The Y axis increment Properties
Public G As Single                     'Gravity
Public V0 As Integer                   'Velocity
Public YM As Single                    'Variable for math calculation
Public Bulx As Double, Buly As Double  'The coordinates at where to fire

'-------------------------------------------------------------------
'This is the sub to call when you want to shoot an object from an
'Object. This is the heart of the trajectory.
'ATTENTION: YOU CAN NOT HAVE ANOTHER CONTROL OR OBJECT NAMED 'FIRE'
Public Sub Fire(source As Object, Ammo As Object, ObjectToDisable As Object, ByVal Angle, ByVal TimeTillExplosion, ByVal Gravity, ByVal Velocity)
'------------------------------------------------------------------
'Source =   Where is the object being fired coming from?
'           What Barrel or tank? It must be a object on the form.
'           Such as a shape or picture, or sprite, etc.
'------------------------------------------------------------------
'Ammo   =   What is being shot/fired? What Object is acting as the
'           Missle or bullet or whatever is being fired/shot?
'           Must be object such as a shape, picture, or sprite, etc
'------------------------------------------------------------------
'ObjectToDisable =  A control on a form such as a button or such
'                   To disable until the ammo hits something or its
'                   Time runs out. I.E. A button, or picture, etc
'-------------------------------------------------------------------
'Angle  =   The angle at which the ammo is being shot. This is a bit
'           Confusing. 0 degrees is completley right. 1.5 is straight
'           up. Thus, 3.14 is straight left. So its in radians. This
'           Can easily be changed by creating a degree formula.
'           Look up  'Radian'  in help. It will give you a formula.
'-------------------------------------------------------------------
'TimeTillExplosion = The amount of time that you want the ammo to fly
'                    Before it stops. (It is Not in seconds be careful
'                    but in a different measurement, not known, in the
'                    sin and cos formulas though.)
'                    For you real mathematicians out there, this really
'                    Just specifies the distance from the source that
'                    You want the ammo to blow up at, But I wrote time
'                    Because I imagine that most people who use this
'                    Will make this a time bomb sort of deal.
'-------------------------------------------------------------------
'Gravity =  How fast you want the ammo to come down. Default is 9.8
'           If you don't want any pull down set '0' as the integer.
'-------------------------------------------------------------------
'Velocity=  How fast do you want the Ammo to move?
'           No unit specified so its not ft/sec or anything.
'-------------------------------------------------------------------
If source = "" Then: Exit Sub       'These are the only needed
If Ammo = "" Then: Exit Sub         'Variables, the others are

Bulx = source.Left + source.Width / 2   'Find the source barrel/gun
Buly = source.Top - 5                   'Coordinates.
                                                '5' can be modified.
Ammo.Left = Bulx                        'Set the Ammo coordinates to
Ammo.Top = Buly                         'The barrel/gun coordinates.

YM = 0                                  'Leave this be.

G = Gravity                             'Sets the gravity Variable.

V0 = Velocity                           'Sets the speed of the Ammo.

ObjectToDisable.Enabled = False         'Disable object specified.

                                        'Calculating...
For T = 0 To TimeTillExplosion Step 0.00004
   'T = 0 means starting the time till it explodes to zero.
            'TimeTillExplosion make the timer count to the explosion.
                              'Step means how close to eachother do
                              'you want the ammo to show. Smaller
                              'number means less jerky. Larger means
                              'more jerky.
    A = Angle                           'Set the Angle.
    'DoEvents   'Add this if you want, it does slow it down though
    DrawBullets Ammo, source            'Draw the Ammo being shot.
Next                                    'Repeat this till timer ends.
ObjectToDisable.Enabled = True                 'Enable the object specified.

'______________________________ <--Actions of Ammo go here
End Sub
'-------------------------------------------------------------------
'This sub draws the ammo.
Sub DrawBullets(Ammo As Object, source As Object)
    X = V0 * T * Cos(A)                 'Figure out the X-Axis proper.
    Y = V0 * T * Sin(A) - G * T ^ 2 / 2 'Fig. the Y Axis properties
                                        'And take into consideration
                                        'The gravity.
    
    If Y > YM Then YM = Y               'Organizes this mess.

    Ammo.Visible = True                 'Show the ammo
    
    'Check below this sub for modifications that you can do to the
    'Ammo.Move function below.
    
    Ammo.Move Bulx + X, Buly - 2 * Y   'SHOOT THE AMMO!!!
    
    'Use the code below if you want to trace your ammo.
    'form1.picture1.PSet (Bulx + X, Buly - 2 * Y)
End Sub
'--------------------------------------------------------------------
        'To make the Ammo.Move move until it hits a certain Y-Axis value
        'then type in front of the Ammo.move Bulx...
        'If Y > (your value) Then:
        'Y = the source.top value.
        'Here is an example of it in action:
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'Sub DrawBullets(Ammo As Object, source As Object)
'X = V0 * T * Cos(A)                 'Figure out the X-Axis proper.
'Y = V0 * T * Sin(A) - G * T ^ 2 / 2 'Fig. the Y Axis properties
                                        'And take into consideration
                                        'The gravity.
    
'If Y > YM Then YM = Y               'Organizes this mess.

'Ammo.Visible = True                 'Show the ammo
    
        'Check below this sub for modifications that you can do to the
        'Ammo.Move function below.
'If Y > -2 Then: Ammo.Move Bulx + X, Buly - 2 * Y    'SHOOT THE AMMO!!!
    
        'Use the code below if you want to trace your ammo.
'form1.picture1.PSet (Bulx + X, Buly - 2 * Y)
'End Sub
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
