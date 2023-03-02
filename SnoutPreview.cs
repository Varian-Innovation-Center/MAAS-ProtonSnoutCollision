using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Windows.Controls;
using System.Windows.Media.Media3D;
using System.Windows.Media;
using System.Windows.Data;
using System.Collections;

[assembly: AssemblyVersion("1.0.0.1")]

namespace VMS.TPS
{
    //GUI class
    public class WNDContent : UserControl
    {
        const double snout_pos_min = 40;
        const double snout_pos_max = 421;
        //snout cover sizes for 3D model
        const double snout_face_zmin = -200;
        const double snout_face_zmax = 200;
        const double snout_face_xmin = -200;
        const double snout_face_xmax = 200;
        const double snout_end_zmin = -250;
        const double snout_end_zmax = 250;
        const double snout_end_xmin = -250;
        const double snout_end_xmax = 250;
        const double snout_depth = 100;

        private Canvas canvas;
        private ComboBox field;
        private PerspectiveCamera camera;
        private Model3DGroup model3D;
        private Label snout_position;
        private Label view_angle;
        private TextBox txt_x;
        private TextBox txt_y;
        private Slider sl_snout_position;
        private Label air_gap;
        private Button btn_calculate;
        private Vector3D plan_snout_position;
        private Vector3D isocenter;
        private Double plan_snout_distance;
        private Double gantry_angle;

        public ScriptContext context { get; set; }

        public WNDContent()
        {
            Grid main_grid = new Grid();

            //Top left section:
            Border top_left = new Border();
            top_left.Width = 300;
            top_left.Height = 70;
            top_left.HorizontalAlignment = HorizontalAlignment.Left;
            top_left.VerticalAlignment = VerticalAlignment.Top;
            top_left.BorderThickness = new Thickness(1,1,2,2);
            top_left.CornerRadius = new CornerRadius(3);
            top_left.BorderBrush = Brushes.Brown;
            top_left.Margin = new Thickness(5, 5, 0, 0);
            main_grid.Children.Add(top_left);

            Grid top_left_grid = new Grid();

            Label lbl_patient = new Label();
            lbl_patient.Content = "Patient:";
            top_left_grid.Children.Add(lbl_patient);

            Label lbl_plan = new Label();
            lbl_plan.Content = "Plan:";
            lbl_plan.Margin = new Thickness(0, 15, 0, 0);
            top_left_grid.Children.Add(lbl_plan);

            Label patient = new Label();
            patient.Name = "patient";
            patient.Margin = new Thickness(50, 0, 0, 0);
            top_left_grid.Children.Add(patient);

            Label plan = new Label();
            plan.Name = "plan";
            plan.Margin = new Thickness(50, 15, 0, 0);
            top_left_grid.Children.Add(plan);

            Label lbl_field = new Label();
            lbl_field.Content = "Select field:";
            lbl_field.Margin = new Thickness(0, 40, 0, 0);
            top_left_grid.Children.Add(lbl_field);

            field = new ComboBox();
            field.Name = "fields";
            field.Height = 20;
            field.Width = 150;
            field.Margin = new Thickness(50, 40, 0, 0);
            field.SelectionChanged += Field_SelectionChanged;
            top_left_grid.Children.Add(field);

            top_left.Child = top_left_grid;

            //Top right section:
            Border top_right = new Border();
            top_right.Width = 300;
            top_right.Height = 70;
            top_right.HorizontalAlignment = HorizontalAlignment.Left;
            top_right.VerticalAlignment = VerticalAlignment.Top;
            top_right.BorderThickness = new Thickness(1, 1, 2, 2);
            top_right.CornerRadius = new CornerRadius(3);
            top_right.BorderBrush = Brushes.Brown;
            top_right.Margin = new Thickness(310, 5, 0, 0);
            main_grid.Children.Add(top_right);

            Grid top_right_grid = new Grid();

            Label lbl_airgap = new Label();
            lbl_airgap.Content = "Measure air gap at resolution:";
            lbl_airgap.Margin = new Thickness(0, 15, 0, 0);
            top_right_grid.Children.Add(lbl_airgap);

            Label lbl_x = new Label();
            lbl_x.Content = "x[mm]:";
            lbl_x.Margin = new Thickness(165, 0, 0, 0);
            top_right_grid.Children.Add(lbl_x);

            Label lbl_y = new Label();
            lbl_y.Content = "y[mm]:";
            lbl_y.Margin = new Thickness(210, 0, 0, 0);
            top_right_grid.Children.Add(lbl_y);

            btn_calculate = new Button();
            btn_calculate.Content = "Calculate";
            btn_calculate.Width = 60;
            btn_calculate.Height = 20;
            btn_calculate.Margin = new Thickness(5, 42, 0, 0);
            btn_calculate.HorizontalAlignment = HorizontalAlignment.Left;
            btn_calculate.VerticalAlignment = VerticalAlignment.Top;
            btn_calculate.Click += Btn_calculate_Click;
            top_right_grid.Children.Add(btn_calculate);

            txt_x = new TextBox();
            txt_x.Width = 35;
            txt_x.Height = 20;
            txt_x.Margin = new Thickness(170, 20, 0, 0);
            txt_x.HorizontalAlignment = HorizontalAlignment.Left;
            txt_x.VerticalAlignment = VerticalAlignment.Top;
            txt_x.PreviewKeyDown += txt_PreviewKeyDown;
            txt_x.TextChanged += txt_TextChanged;
            txt_x.Text = "10";
            top_right_grid.Children.Add(txt_x);

            txt_y = new TextBox();
            txt_y.Width = 35;
            txt_y.Height = 20;
            txt_y.Margin = new Thickness(215, 20, 0, 0);
            txt_y.HorizontalAlignment = HorizontalAlignment.Left;
            txt_y.VerticalAlignment = VerticalAlignment.Top;
            txt_y.PreviewKeyDown += txt_PreviewKeyDown;
            txt_y.TextChanged += txt_TextChanged; ;
            txt_y.Text = "10";
            top_right_grid.Children.Add(txt_y);

            air_gap = new Label();
            air_gap.Content = "Calculated air gap =";
            air_gap.Margin = new Thickness(100, 40, 0, 0);
            top_right_grid.Children.Add(air_gap);

            top_right.Child = top_right_grid;

            //Canvas controls
            Label lbl_view_angle = new Label();
            lbl_view_angle.Content = "View angle[deg]: ";
            lbl_view_angle.Margin = new Thickness(5, 80, 0, 0);
            main_grid.Children.Add(lbl_view_angle);

            view_angle = new Label();
            view_angle.Margin = new Thickness(95, 80, 0, 0);
            view_angle.Content = 0;
            main_grid.Children.Add(view_angle);

            Slider sl_view_angle = new Slider();
            sl_view_angle.Width = 450;
            sl_view_angle.Margin = new Thickness(160, 85, 0, 0);
            sl_view_angle.HorizontalAlignment = HorizontalAlignment.Left;
            sl_view_angle.Minimum = -180;
            sl_view_angle.Maximum = 180;
            sl_view_angle.Value = 0;
            sl_view_angle.ValueChanged += sl_view_angle_ValueChanged;
            main_grid.Children.Add(sl_view_angle);

            Label lbl_snout_position = new Label();
            lbl_snout_position.Content = "Snout position[cm]: ";
            lbl_snout_position.Margin = new Thickness(5, 100, 0, 0);
            main_grid.Children.Add(lbl_snout_position);

            snout_position = new Label();
            snout_position.Name = "snout_position_value";
            snout_position.Margin = new Thickness(110, 100, 0, 0);
            main_grid.Children.Add(snout_position);

            sl_snout_position = new Slider();
            sl_snout_position.Name = "snout_position";
            sl_snout_position.Width = 450;
            sl_snout_position.Margin = new Thickness(160, 105, 0, 0);
            sl_snout_position.HorizontalAlignment = HorizontalAlignment.Left;
            sl_snout_position.Minimum = snout_pos_min;
            sl_snout_position.Maximum = snout_pos_max;
            sl_snout_position.Value = 0;
            sl_snout_position.ValueChanged += sl_snout_position_ValueChanged;
            main_grid.Children.Add(sl_snout_position);

            canvas = new Canvas();
            canvas.Margin = new Thickness(0, 130, 0, 0);
            canvas.Background = Brushes.LightSkyBlue;
            canvas.MouseWheel += Canvas_MouseWheel;
            main_grid.Children.Add(canvas);

            Button btn_help = new Button();
            btn_help.Width = 20;
            btn_help.Height = 20;
            btn_help.Content = "?";
            btn_help.Click += Btn_help_Click;
            btn_help.Background = new SolidColorBrush(Color.FromArgb(128, 221, 221, 221));
            canvas.Children.Add(btn_help);
            Canvas.SetRight(btn_help, 0);

            this.Content = main_grid;
        }

        private void Btn_help_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("This script visualizes a snout for proton machine. \r\n\r\n" +
                "It allows moving the snout and calculating air gap.\r\n\r\n" + 
                "Air gap is calcuated using ray tracing from the snout down to the body. Raytracing can be done at a custom resolution.\r\n\r\n"+
                "Note: small value may require a few seconds to calculate depending on snout size; couch rotation is ignored and visualization will be incorrect in the current implementation");
        }

        private void txt_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox txt = sender as TextBox;
            if (txt.Text.Length == 0)
            {
                (sender as TextBox).Background = Brushes.OrangeRed;
                btn_calculate.IsEnabled = false;
            }
            else
            {
                Double input;
                Boolean IsNumber = Double.TryParse(txt.Text, out input);
                if (IsNumber)
                {
                    if ((input > 0) && (input <= 50))
                    {
                        (sender as TextBox).Background = Brushes.White;
                        btn_calculate.IsEnabled = true;
                    }
                    else
                    {
                        (sender as TextBox).Background = Brushes.OrangeRed;
                        btn_calculate.IsEnabled = false;
                    }
                }
                else
                {
                    (sender as TextBox).Background = Brushes.OrangeRed;
                    btn_calculate.IsEnabled = false;
                }
            }
        }

        //Calculate airgap at the given offset
        private void Btn_calculate_Click(object sender, RoutedEventArgs e)
        {
            //Get reference to body outline
            IonPlanSetup active_proton_plan = context.IonPlanSetup;
            StructureSet sset = active_proton_plan.StructureSet;
            Structure body = sset.Structures.Where(s => s.DicomType == "EXTERNAL").First();

            Double resolution_x = Convert.ToDouble(txt_x.Text);
            Double resolution_y = Convert.ToDouble(txt_y.Text);

            int x_counter = 0;
            Double air_gap_value = Double.MaxValue;
            VVector closest_collision_point = new VVector(0,0,0);
            VVector offset_at_snout = new VVector(0, 0, 0);
            Boolean in_collision = false;
            while (snout_face_xmin + x_counter * resolution_x <= snout_face_xmax)
            {
                int y_counter = 0;
                while (snout_face_zmin + y_counter * resolution_y <= snout_face_zmax)
                {
                    Vector3D shift = new Vector3D(snout_face_xmin + x_counter * resolution_x, 0, snout_face_zmin + y_counter * resolution_y);

                    Double current_snout_position = sl_snout_position.Value;
                    Vector3D iso_shifted = Vector3D.Add(isocenter, shift);
                    Vector3D snout_center_shifted = Vector3D.Add(new Vector3D(0, -current_snout_position, 0), shift);

                    //gantry rotation matrix
                    Matrix3D matrix3D = Matrix3D.Identity;
                    Quaternion rot = new Quaternion(new Vector3D(0, 0, 1), gantry_angle);
                    matrix3D.Rotate(rot);

                    //apply gantry angle
                    iso_shifted = matrix3D.Transform(iso_shifted);
                    snout_center_shifted = matrix3D.Transform(snout_center_shifted);

                    snout_center_shifted = Vector3D.Add(snout_center_shifted, isocenter);

                    VVector end = new VVector(iso_shifted.X, iso_shifted.Y, iso_shifted.Z);
                    VVector start = new VVector(snout_center_shifted.X, snout_center_shifted.Y, snout_center_shifted.Z);

                    BitArray array = new BitArray(10 * (int)current_snout_position);

                    SegmentProfile profile = body.GetSegmentProfile(start, end, array);

                    int collision_index = 0;
                    Boolean found = false;
                    foreach (SegmentProfilePoint p in profile)
                    {
                        if (p.Value == true)
                        {
                            found = true;
                            break;
                        }
                        collision_index++;
                    }

                    if (found)
                    {
                        if (collision_index == 0)
                        {
                            in_collision = true;
                        }
                        else
                        {
                            VVector collision_point = new VVector(profile[collision_index].Position.x, profile[collision_index].Position.y, profile[collision_index].Position.z);

                            Double current_air_gap_value = (collision_point - start).Length;

                            if (current_air_gap_value < air_gap_value)
                            {
                                air_gap_value = current_air_gap_value;

                                closest_collision_point = collision_point;
                                offset_at_snout = start;
                            }
                        }

                    }
                    else
                    {
                        //not found
                    }
                    y_counter++;
                }
                x_counter++;
            }

            if (in_collision)
            {
                air_gap.Content = "Calculated air gap = IN COLLISION";
            }
            else
            {
                Line3D line = new Line3D(new Point3D(closest_collision_point.x, closest_collision_point.y, closest_collision_point.z),
                    new Point3D(offset_at_snout.x, offset_at_snout.y, offset_at_snout.z), Brushes.Red, 3);
                model3D.Children.Add(line.GeometryModel3D);

                air_gap.Content = "Calculated air gap = " + air_gap_value.ToString("###.0") + " mm";
            }
 
        }

        private void Field_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (model3D != null)
            {
                Initiate3DView();
            }
        }

        private void txt_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (((e.Key >= System.Windows.Input.Key.D0) && (e.Key <= System.Windows.Input.Key.D9)) || ((e.Key >= System.Windows.Input.Key.NumPad0) && (e.Key <= System.Windows.Input.Key.NumPad9)) ||
                (e.Key == System.Windows.Input.Key.Back) || (e.Key == System.Windows.Input.Key.Delete))
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        //to zoom in and out the 3D model with mouse wheel
        private void Canvas_MouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            Point3D point = new Point3D(camera.Position.X, camera.Position.Y, camera.Position.Z);
            Matrix3D scalematrix = Matrix3D.Identity;
            if (e.Delta > 0)
            {
                scalematrix.Scale(new Vector3D(1.1, 1.1, 1.1));
            }
            else
            {
                scalematrix.Scale(new Vector3D(0.9, 0.9, 0.9));
            }
            point = Point3D.Multiply(point, scalematrix);
            camera.Position = point;
        }

        //to rotate view with slider
        private void sl_view_angle_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            view_angle.Content = (sender as Slider).Value.ToString("000.0");

            if (model3D != null)
            {
                model3D.Transform = new RotateTransform3D(new AxisAngleRotation3D(new Vector3D(0, 0, 1), e.NewValue), 0, 0, 0);
            }
        }

        //to move the snout
        private void sl_snout_position_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            Slider slider = sender as Slider;

            if (model3D != null)
            {
                //Erase previously calculated airgap:
                if (model3D.Children.Count == 6) //includes calculated gap which is the last child
                {
                    model3D.Children.RemoveAt(5);
                    air_gap.Content = "Calculated air gap = ";
                }

                snout_position.Content = (e.NewValue / 10).ToString("##.0");

                //getting reference to snout mesh, snout model is the second child of the model3D
                Model3DGroup snout_model = model3D.Children[2] as Model3DGroup;

                Vector3D snout_axis = Vector3D.Subtract(plan_snout_position, isocenter);

                snout_model.Transform = new TranslateTransform3D(Vector3D.Multiply(e.NewValue / plan_snout_distance - 1, snout_axis));
            }
        }

        public void Initiate3DView()
        {
            //collecting required patient information
            IonPlanSetup active_proton_plan = context.IonPlanSetup;
            IEnumerable<Beam> fields = active_proton_plan.Beams;

            StructureSet sset = active_proton_plan.StructureSet;
            Structure body = sset.Structures.Where(s => s.DicomType == "EXTERNAL").First();

            //creating visual model
            ModelVisual3D modelvisual = new ModelVisual3D();
            model3D = new Model3DGroup();

            modelvisual.Content = model3D;

            Viewport3D viewport = new Viewport3D();
            Canvas.SetLeft(viewport, 0);
            Canvas.SetTop(viewport, 0);
            if (canvas.Children.Count == 2)
            {
            //When creating 3D view canvas has only 1 child. After that it will have 2, second one  being the 3 view
            canvas.Children.RemoveAt(1);
            }
            canvas.Children.Add(viewport);
            Binding binding_w = new Binding();
            binding_w.Path = new PropertyPath(Canvas.ActualWidthProperty);
            binding_w.Source = canvas;
            viewport.SetBinding(Viewport3D.WidthProperty, binding_w);
            Binding binding_h = new Binding();
            binding_h.Path = new PropertyPath(Canvas.ActualHeightProperty);
            binding_h.Source = canvas;
            viewport.SetBinding(Viewport3D.HeightProperty, binding_h);

            //assigning model visual to view port 3D of the canvas on the main window
            viewport.Children.Add(modelvisual);

            //adding patient body 3D in the model
            GeometryModel3D patientmodel = new GeometryModel3D();
            patientmodel.Geometry = body.MeshGeometry;
            DiffuseMaterial dm = new DiffuseMaterial();
            dm.Brush = new SolidColorBrush(Color.FromScRgb(1, 0.8f, 0.7176470588235294f, 0.4784313725490196f));
            patientmodel.Material = dm;
            model3D.Children.Add(patientmodel);

            //creating snout mesh at well know geometry when gantry at 0  --------------------------------------------

            //read required information from the plan, for a field selected in drop down
            gantry_angle = fields.ElementAt(field.SelectedIndex).ControlPoints[0].GantryAngle;
            VVector v_isocenter = fields.ElementAt(field.SelectedIndex).IsocenterPosition;
            plan_snout_distance = (fields.ElementAt(field.SelectedIndex) as IonBeam).SnoutPosition * 10; //converting to mm

            isocenter = new Vector3D(v_isocenter.x, v_isocenter.y, v_isocenter.z);

            //snout face center with gantry at 0            
            Vector3D snout_center_pos = new Vector3D(0, -plan_snout_distance, 0);

            //Calculating corners of snout cover

            //front/proximal face
            Vector3D bottomleftfront = Vector3D.Add(snout_center_pos, new Vector3D(snout_face_xmin, 0, snout_face_zmin)); //assuming front face is square +/- 20cm around center
            Vector3D bottomrightfront = Vector3D.Add(snout_center_pos, new Vector3D(snout_face_xmax, 0, snout_face_zmin));
            Vector3D topleftfront = Vector3D.Add(snout_center_pos, new Vector3D(snout_face_xmin, 0, snout_face_zmax));
            Vector3D toprightfront = Vector3D.Add(snout_center_pos, new Vector3D(snout_face_xmax, 0, snout_face_zmax));
            //back/distal end, assuming conical shape, size increases distally
            Vector3D bottomleftback = Vector3D.Add(snout_center_pos, new Vector3D(snout_end_xmin, -snout_depth, snout_end_zmin)); //assuming back side is square +/- 25cm around center, thickness 10cm
            Vector3D bottomrightback = Vector3D.Add(snout_center_pos, new Vector3D(snout_end_xmax, -snout_depth, snout_end_zmin));
            Vector3D topleftback = Vector3D.Add(snout_center_pos, new Vector3D(snout_end_xmin, -snout_depth, snout_end_zmax));
            Vector3D toprightback = Vector3D.Add(snout_center_pos, new Vector3D(snout_end_xmax, -snout_depth, snout_end_zmax));

            //gantry rotation matrix
            Matrix3D matrix3D = Matrix3D.Identity;
            Quaternion rot = new Quaternion(new Vector3D(0, 0, 1), gantry_angle);
            matrix3D.Rotate(rot);

            //Rotating snout (all its corners) by gantry angle
            bottomleftfront = matrix3D.Transform(bottomleftfront);
            bottomrightfront = matrix3D.Transform(bottomrightfront);
            topleftfront = matrix3D.Transform(topleftfront);
            toprightfront = matrix3D.Transform(toprightfront);
            //back/distal end, assuming conical shape, size increases distally
            bottomleftback = matrix3D.Transform(bottomleftback);
            bottomrightback = matrix3D.Transform(bottomrightback);
            topleftback = matrix3D.Transform(topleftback);
            toprightback = matrix3D.Transform(toprightback);

            plan_snout_position = matrix3D.Transform(snout_center_pos);

            //Shifting by isocenter:
            plan_snout_position = Vector3D.Add(plan_snout_position, isocenter);

            bottomleftfront = Vector3D.Add(bottomleftfront, isocenter);
            bottomrightfront = Vector3D.Add(bottomrightfront, isocenter);
            topleftfront = Vector3D.Add(topleftfront, isocenter);
            toprightfront = Vector3D.Add(toprightfront, isocenter);
            bottomleftback = Vector3D.Add(bottomleftback, isocenter);
            bottomrightback = Vector3D.Add(bottomrightback, isocenter);
            topleftback = Vector3D.Add(topleftback, isocenter);
            toprightback = Vector3D.Add(toprightback, isocenter);

            //defining snout mesh for rotated snout
            Model3DGroup snout = new Model3DGroup();
            GeometryModel3D snout_model = new GeometryModel3D();
            MeshGeometry3D snout_mesh = new MeshGeometry3D();
            Point3D Identity = new Point3D(0, 0, 0); //to convert vectors to points for position definitions
            snout_mesh.Positions = new Point3DCollection() { Point3D.Add(Identity, bottomleftfront), Point3D.Add(Identity, bottomrightfront), Point3D.Add(Identity, topleftfront), Point3D.Add(Identity, toprightfront),
             Point3D.Add(Identity, bottomleftback), Point3D.Add(Identity, bottomrightback), Point3D.Add(Identity, topleftback), Point3D.Add(Identity, toprightback)};
            //back traingles not includes, just a hollow cover
            snout_mesh.TriangleIndices = new Int32Collection() { 0, 2, 1, 1, 2, 3, 0, 5, 4, 0, 1, 5, 1, 7, 5, 1, 3, 7, 3, 6, 7, 3, 2, 6, 2, 4, 6, 2, 0, 4 };
            snout_model.Geometry = snout_mesh;
            DiffuseMaterial snout_material = new DiffuseMaterial();
            snout_material.Brush = new SolidColorBrush(Colors.LightGray);
            snout_model.Material = snout_material;
            //inner side of snout
            DiffuseMaterial snout_back_material = new DiffuseMaterial();
            snout_back_material.Brush = new SolidColorBrush(Colors.Red);
            snout_model.BackMaterial = snout_back_material;
            snout.Children.Add(snout_model);

            //creating frame
            Line3D front_bottom = new Line3D(Point3D.Add(Identity, bottomleftfront), Point3D.Add(Identity, bottomrightfront), Brushes.Black, 2);
            snout.Children.Add(front_bottom.GeometryModel3D);
            Line3D front_top = new Line3D(Point3D.Add(Identity, topleftfront), Point3D.Add(Identity, toprightfront), Brushes.Black, 2);
            snout.Children.Add(front_top.GeometryModel3D);
            Line3D front_right = new Line3D(Point3D.Add(Identity, bottomrightfront), Point3D.Add(Identity, toprightfront), Brushes.Black, 2);
            snout.Children.Add(front_right.GeometryModel3D);
            Line3D front_left = new Line3D(Point3D.Add(Identity, bottomleftfront), Point3D.Add(Identity, topleftfront), Brushes.Black, 2);
            snout.Children.Add(front_left.GeometryModel3D);
            Vector3D snout_parked = Vector3D.Multiply(Vector3D.Divide(Vector3D.Subtract(plan_snout_position, isocenter), plan_snout_distance), snout_pos_max);
            Line3D CAX = new Line3D(Point3D.Add(Identity, Vector3D.Add(snout_parked,isocenter)), new Point3D(isocenter.X, isocenter.Y, isocenter.Z), Brushes.YellowGreen, 1);
            model3D.Children.Add(CAX.GeometryModel3D);

            model3D.Children.Add(snout);

            //adding lights 
            DirectionalLight light1 = new DirectionalLight(Colors.WhiteSmoke, new Vector3D(0, 600, 0));
            model3D.Children.Add(light1);
            DirectionalLight light2 = new DirectionalLight(Colors.WhiteSmoke, new Vector3D(0, -600, 0));
            model3D.Children.Add(light2);

            //adding camera (facing the patient at 100cm distance
            camera = new PerspectiveCamera ();
            camera.Position = new Point3D(0, -1500, 0);
            camera.LookDirection = new Vector3D(0, 1500, 0);
            camera.UpDirection = new Vector3D(0, 0, 1);   

            viewport.Camera = camera;
        }

    }

    //Class representing 3D line
    public class Line3D
    {
        private GeometryModel3D geoModel3D;
        private Point3D start;
        private Point3D end;

        Vector3D norm_to_line1;
        Vector3D norm_to_line2;

        public GeometryModel3D GeometryModel3D
        { get 
            {   
                return geoModel3D;   
            }
        }

        private void GenerateVertices(Point3DCollection points)
        {
            points.Add(Point3D.Add(start,norm_to_line1));
            points.Add(Point3D.Add(start, norm_to_line2));
            points.Add(Point3D.Add(end, norm_to_line1));
            points.Add(Point3D.Add(end, norm_to_line2));
            points.Add(Point3D.Add(start, -norm_to_line1));
            points.Add(Point3D.Add(start, -norm_to_line2));
            points.Add(Point3D.Add(end, -norm_to_line1));
            points.Add(Point3D.Add(end, -norm_to_line2));
        }

        public Line3D(Point3D start, Point3D end, Brush brush, Double line_thickness)
        {
            this.end = end;
            this.start = start;

            Vector3D line = Point3D.Subtract(end, start);

            if ((Math.Abs(line.X) > Math.Abs(line.Y)) && (Math.Abs(line.X) > Math.Abs(line.Z)))
            {
                norm_to_line1 = new Vector3D( -(line.Y + line.Z)/line.X , 1, 1);
            }

            if ((Math.Abs(line.Y) > Math.Abs(line.X)) && (Math.Abs(line.Y) > Math.Abs(line.Z)))
            {
                norm_to_line1 = new Vector3D(1,-(line.X + line.Z) / line.Y, 1);
            }

            if ((Math.Abs(line.Z) > Math.Abs(line.X)) && (Math.Abs(line.Z) > Math.Abs(line.Y)))
            {
                norm_to_line1 = new Vector3D(1, 1, -(line.X + line.Y) / line.Z);
            }

            norm_to_line2 = Vector3D.CrossProduct(line, norm_to_line1);

            //Normalization:
            norm_to_line1 = Vector3D.Divide(norm_to_line1, norm_to_line1.Length);
            norm_to_line2 = Vector3D.Divide(norm_to_line2, norm_to_line2.Length);
            norm_to_line1 = Vector3D.Multiply(norm_to_line1, line_thickness);
            norm_to_line2 = Vector3D.Multiply(norm_to_line2, line_thickness);

            geoModel3D = new GeometryModel3D();
            DiffuseMaterial line_material = new DiffuseMaterial(brush);
            line_material.AmbientColor = Colors.Blue;
            geoModel3D.Material = line_material;

            MeshGeometry3D line_mesh = new MeshGeometry3D();
            GenerateVertices(line_mesh.Positions);
            line_mesh.TriangleIndices = new Int32Collection() {0,1,2  ,2,1,3  ,5,7,4  ,7,6,4  ,1,4,6  ,6,3,1  ,5,2,7  ,0,2,5  ,7,2,3  ,7,3,6  ,4,0,5  ,4,1,0};
            geoModel3D.Geometry = line_mesh;

        }
         
    }

    public class Script
    {
        public Script()
        {
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context, System.Windows.Window window, ScriptEnvironment environment)
        {
            window.Activated += Window_Activated;

            if (context.Patient == null)
            {
                MessageBox.Show("There is no patient opened. Please open patient and a proton plan.");
                return;
            }

            if (context.IonPlanSetup == null)
            {
                MessageBox.Show("There are no proton plans opened. Please open a proton plan.");
                return;
            }

            WNDContent wnd = new WNDContent();
            wnd.context = context;
            window.Content = wnd;
            window.MinWidth = 630;
            window.MinHeight = 800;
            window.Width = 630;
            window.Height = 800;

            //Initialize GUI
            Label lbl = LogicalTreeHelper.FindLogicalNode(window, "patient") as Label;
            lbl.Content = context.Patient.FirstName + " " + context.Patient.LastName + " (ID:" + context.Patient.Id +  ")";
            
            lbl = LogicalTreeHelper.FindLogicalNode(window, "plan") as Label;
            lbl.Content = context.Patient.FirstName + " " + context.IonPlanSetup;

            Double snout_distance = context.IonPlanSetup.IonBeams.ElementAt(0).SnoutPosition;
            lbl = LogicalTreeHelper.FindLogicalNode(window, "snout_position_value") as Label;
            lbl.Content = snout_distance.ToString("##.0");

            Slider sl = LogicalTreeHelper.FindLogicalNode(window, "snout_position") as Slider;
            sl.Value = snout_distance * 10;

            ComboBox cb = LogicalTreeHelper.FindLogicalNode(window, "fields") as ComboBox;
            foreach (IonBeam beam in context.IonPlanSetup.IonBeams)
            {
                cb.Items.Add(beam.Id);
            }
            cb.SelectedIndex = 0;

            wnd.Initiate3DView();
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            //Closing empty window if patient or plan not opened;
            Window window = sender as Window;
            if (window.Content == null)
            {
                window.Close();
            }
        }

    }
}
