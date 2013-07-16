SourceScanner
----

This little utility written in VB builds a .DEF/.ENT files used 
by GtkRadiant directly from mod's source (finding /*QUAKED entries)

Feature list
------

- Extracting of /*QUAKED entity definitions from source files and placing it in entities.def.
- Extracting of /*SS definitions into entities.ent (a GtkRadiant 1.5-introduced XML entity definitions) 
- Using defines and templates that readed from external ini file which should be located in: [destination folder]\ssconfig.ini or [destination folder]\..\[destination folder].ini

Using /*SS entries
------

Source:

/*SS 
form =quaked;
lang =english;
class=func_train;
name =Train;
desc =Trains are moving platforms that players can ride. 
Trains must have "origin" brush inside it to know what point to use for corners placement. Otherwise XYZ min corner will be used. 

Use "path_corner" entities for train movement spots, when train will reach them, it will wait path_corner's "wait" seconds (-1 to wait forever), and then change it's targets to path_corner's ones and continue move. Also "speed" of movement could be set to override train default speed, it's determined from target path_corner when train move starts.

Train spawns at the first target it is pointing at (or saves it's position if no target is set), if "wait" is set, train will travel to it's starting target with "wait" seconds delay. Value of -1 means forever delay.

If train is the target of a button or trigger, it will not begin moving until activated.

Targetparms's to call with:
- train:go - causes train to move to next target (thats useful when targets have wait -1)
- train:goto [targetname] - causes train to change target and start move to it from current position
- train:stop - stops train
- train:setspeed - stops current train, changes speed and let train continue;
key  =speed\integer\20\Move speed\Train move speed in game units per second;
key  =wait\integer\0\Start delay\Delay in seconds before start move;		
key  =dmg\integer\0\Damage\Amount damage train is blocked (or touched if has SPIKES spawnflag) by someone;		
key  =lip\integer\4\Spikes lip\Spikes damagezone extends train size by this amount of units, use with SPIKES spawnflag;		
flag =1\AUTOSTOP\AUTOSTOP\Causes train to stop on each corner it reachs;
flag =2\SPIKES\SPIKES\Causes train to damage player if touches it, defaults damage to 5;
flag =4\JUMP_OVER\JUMP_OVER\Allow jump over this barrier in lupon form;
group=bmodel;
group=bmodel_func;
group=gamemodel;
group=entity;
*/

Extracted as:

<group name="func_train" color="0.3 0.3 1"> Trains are moving platforms that players can ride.
 Trains must have "origin" brush inside it to know what point to use for corners placement. Otherwise XYZ min corner will be used.
 Use "path_corner" entities for train movement spots, when train will reach them, it will wait path_corner's "wait" seconds (-1 to wait forever), and then change it's targets to path_corner's ones and continue move. Also "speed" of movement could be set to override train default speed, it's determined from target path_corner when train move starts.
 Train spawns at the first target it is pointing at (or saves it's position if no target is set), if "wait" is set, train will travel to it's starting target with "wait" seconds delay. Value of -1 means forever delay.
 If train is the target of a button or trigger, it will not begin moving until activated.
 Targetparms's to call with:
 - train:go - causes train to move to next target (thats useful when targets have wait -1)
 - train:goto [targetname] - causes train to change target and start move to it from current position
 - train:stop - stops train
 - train:setspeed - stops current train, changes speed and let train continue

 --- Keys --- 
<integer key="speed" name="Move speed">Train move speed in game units per second. Default is 20.</integer>
<integer key="wait" name="Start delay">Delay in seconds before start move. Default is 0.</integer>
<integer key="dmg" name="Damage">Amount damage train is blocked (or touched if has SPIKES spawnflag) by someone. Default is 0.</integer>
<integer key="lip" name="Spikes lip">Spikes damagezone extends train size by this amount of units, use with SPIKES spawnflag. Default is 4.</integer>
<real key="_ls" name="Lightmap scale">Scale lightmap sample size on this model, 0.5 is 2 times better lightmaps;.</real>
<real key="_cs" name="Lightmap shadow cast group">Lightmap shadow cast group;.</real>
<real key="_rs" name="Lightmap shadow receive group">Lightmap shadow cast group;.</real>
<real key="_vp" name="Vertical texture projection">Vertical texture projection;.</real>
<real key="_np" name="Smooth normals">Smooth normals;.</real>
<real key="_uvs" name="UV Scale">Scale UV by this value, 0.5 is 2x larger texture;.</real>
<boolean key="_pm" name="Patch meta">Patch meta;.</boolean>
<integer key="_ps" name="Patch subdivisions">Patch sibdisision tolerance. Lesser value = more patch triangles generated;. Default is 8.</integer>
<boolean key="_ns" name="Disable BSP">Don not generate collision surfaces on this model (does not apply to shader-forced BIH collisions).</boolean>
<string key="classname" name="Entity class">Entity class;.</string>
<string key="sid" name="Entity SaveID">SaveID of entity;.</string>

 --- Flags --- 
<flag key="AUTOSTOP" name="AUTOSTOP" bit="0">Causes train to stop on each corner it reachs</flag>
<flag key="SPIKES" name="SPIKES" bit="1">Causes train to damage player if touches it, defaults damage to 5</flag>
<flag key="JUMP_OVER" name="JUMP_OVER" bit="3">Allow jump over this barrier in lupon form</flag>
<flag key="IGNORE_LIGHTMAPS" name="Ignore lightmaps" bit="7">This spawnflag forces lightgrid sampling to be used on model, light will be sampled at model origin</flag>

</group>

Using of /*QUAKED entries
------

Source:

/*QUAKED path_corner (0.5 0.3 0) (-8 -8 -8) (8 8 8) ?
Monsters will continue walking towards the next target corner.
*/

Extracted as (no different):

/*QUAKED path_corner (0.5 0.3 0) (-8 -8 -8) (8 8 8) ?
 Monsters will continue walking towards the next target corner.
*/

Using of /*GTKRAD15 entries
------

Source:

/*GTKRAD15
<group name="func_button" color="0 .5 .8">
When a button is touched, it moves some distance in the 
direction of its angle, 
triggers all of its targets, waits some time, then returns 
to its original 
position where it can be triggered again.
<direction key="angle" name="Move Direction">Determines the 
opening direction</direction>
<target key="target" name="Target">All entities with a matching 
targetname will be used</target>
<target key="killtarget" name="Kill Target">Kills the target 
when fired</target>
<real key="speed" name="Move Speed">Default is 40</real>
<real key="wait" name="Return Delay">Default is 1 
(-1 = never return) - does not apply to doors that open 
with a key</real>
<real key="lip" name="Lip">Override the default 4 pixel 
lip remaining at end of move</real>
<boolean key="health" name="Shootable">If set, the button 
must be killed instead of touched</boolean>
<integer key="sounds" name="Sounds">
 0 = steam metal
 1 = wooden clunk
 2 = metallic click
 3 = in-out
</integer>
*/

Extracted as (no difference):

<group name="func_button" color="0 .5 .8">
When a button is touched, it moves some distance in the 
direction of its angle, 
triggers all of its targets, waits some time, then returns 
to its original 
position where it can be triggered again.
<direction key="angle" name="Move Direction">Determines the 
opening direction</direction>
<target key="target" name="Target">All entities with a 
matching targetname will be used</target>
<target key="killtarget" name="Kill Target">Kills the target 
when fired</target>
<real key="speed" name="Move Speed">Default is 40</real>
<real key="wait" name="Return Delay">Default is 1 
(-1 = never return) - does not apply to doors that open 
with a key</real>
<real key="lip" name="Lip">Override the default 4 pixel lip 
remaining at end of move</real>
<boolean key="health" name="Shootable">If set, the button 
must be killed instead of touched</boolean>
<integer key="sounds" name="Sounds">
 0 = steam metal
 1 = wooden clunk
 2 = metallic click
 3 = in-out
</integer>

--------------------------------------------------------------------------------
 Version History + Changelog (Reverse Chronological Order)
--------------------------------------------------------------------------------

1.0 (16-07-2013)
------
- initial commit (borrowed from archives)