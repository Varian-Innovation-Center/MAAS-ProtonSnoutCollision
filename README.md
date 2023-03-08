# MAAS-ProtonSnoutCollision

Snout Preview (SnoutPreview.cs) is Eclipse script written to visualize and detect collisions between a snout of proton treatment machine and patient.

Background: Eclipse limits collision detection to area that is defined as ‘Maximum Field Size’. This does not include the cover of the snout that is usually significantly larger than the ‘Maximum Field Size’. Therefore, it can happen that when a proton made is made and opened to treat a patient, the treatment is not possible due to a collision between snout and patient that went undetected at the Eclipse.

Solution/Mitigation: The script allows visualizing snout in its full size and shows the snout in a 3D view together with patient BODY outline. For proton plans, the BODY outline usually includes also treatment table so visualization includes the treatment table as well. In the script the snout is modelled with 8 points, 4 for the face and 4 for the back side. The shape is kept hollow so it can be seen when the BODY protrudes into the snout cover. The outside color is ‘silver’, the inside surface of the snout is red. More points can be added in the script to model the snout at a desired complexity. The visualization is done per field (user has the option to choose a field from the plan) and considers the field’s gantry angle and current snout position for the field which will be used as the initial snout position for the 3D view.

The script allows rotating the 3D view with the snout and the patient, and the snout can be moved with a Slider control further away or closer to the patient to estimate optimal snout position for the plan. The snout position is displayed in the dialog and the number can be used with the treatment plan. The script also allows calculating an air gap (distance between BODY and face of the snout). The air gap is then visualized with a red line in 3D view which shows the area where the collision is most likely to happen.

Snout Preview is a true script, just a cs file and can be easily modified and enhanced.
