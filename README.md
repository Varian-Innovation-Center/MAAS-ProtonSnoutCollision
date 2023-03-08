# MAAS-ProtonSnoutCollision

![image](https://user-images.githubusercontent.com/78000769/223833151-e04962c4-a286-4490-aa73-e86185d0b85b.png)

Snout Preview (SnoutPreview.cs) is an Eclipse script written to visualize and detect collisions between the snout of a proton treatment machine and patient.

Background: Eclipse limits collision detection to area that is defined as ‘Maximum Field Size’. This does not include the cover of the snout that is usually significantly larger than the ‘Maximum Field Size’. Therefore, it can happen that when a proton field is made and opened to treat a patient, the treatment is not possible due to a collision between snout and patient that went undetected in the Eclipse client software.

Solution/Mitigation: The script allows visualizing snout in its full size and shows the snout in a 3D view together with patient BODY outline. For proton plans, the BODY outline includes the treatment table so visualization includes the treatment table as well. In the script the snout is modelled with 8 points, 4 for the face and 4 for the back side. The shape is kept hollow so it can be seen when the BODY protrudes into the snout cover. The outside color is ‘silver’, the inside surface of the snout is red. More points can be added in the script to model the snout at a desired complexity. The visualization is done per field (user has the option to choose a field from the plan) and considers the field’s gantry angle and current snout position for the field which will be used as the initial snout position for the 3D view.

The script allows rotating the 3D view with the snout and the patient, and the snout can be moved with a Slider control further away or closer to the patient to estimate optimal snout position for the plan. The snout position is displayed in the dialog and the number can be used with the treatment plan. The script also allows calculating an air gap (distance between BODY and face of the snout). The air gap is then visualized with a red line in 3D view which shows the area where the collision is most likely to happen.

Snout Preview is a true script, just a non-compiled cs file and can be easily modified and enhanced.

Note: Eclipse version 16 required
https://github.com/Varian-Innovation-Center/MAAS-ProtonSnoutCollision/blob/52d9df2878d536bf3cbb50f93d7a1b52fb06f81b/SnoutPreview.cs#L23

Tip: look for USER MODIFIABLE sections of the code to turn on/off "not yet validated" warning or licence agreement popup
https://github.com/Varian-Innovation-Center/MAAS-ProtonSnoutCollision/blob/4a4f6194984976cc76f2afdc5d08b49120bc3915/SnoutPreview.cs#L881

Orginal code by Peter Klenovsky
