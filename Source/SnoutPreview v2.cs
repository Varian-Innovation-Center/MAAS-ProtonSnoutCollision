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
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Data;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows.Navigation;
using System.Globalization;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
//using System.Reflection.Emit;

[assembly: AssemblyVersion("1.0.0.1")]

// Caleb Notes:
// Needs an IonBeam plan


namespace VMS.TPS
{
    //Snout
    public class Snout
    {

        const double snout_pos_min = 0;
        const double snout_pos_max = 421;
        //snout cover sizes for 3D model
        const double snout_face_zmin = -350;
        const double snout_face_zmax = 350;
        const double snout_face_xmin = -250;
        const double snout_face_xmax = 250;
        const double snout_end_zmin = -400;
        const double snout_end_zmax = 400;
        const double snout_end_xmin = -300;
        const double snout_end_xmax = 300;
        const double snout_depth = 100;

        private double current_snout_distance;
        private double planned_snout_distance;
        private Vector3D planned_snout_position;
        private double air_gap = Double.NaN;
        private Point3D air_gap_start;
        private Point3D air_gap_end;
        private double gantry_angle;
        private double couch_rotation;
        private Vector3D isocenter;
        private Model3DGroup snout_geometry;
        private MeshGeometry3D snout_mesh;
        private bool in_collision = false;
        private bool air_gap_not_found = true;

        public bool In_Collision
        { 
            get { return in_collision; }
        }

        public bool Air_Gap_Not_Found
        {
            get { return air_gap_not_found; }
        }

        public double Snout_Distance
        {
            get { return current_snout_distance; }
        }

        public Vector3D Planned_Snout_Position
        {
            get { return planned_snout_position; }
        }

        public Model3DGroup Geometry
        { 
            get { return snout_geometry; }
        }

        public double Air_Gap
        { 
            get { return air_gap; }
        }

        public Point3D Air_Gap_Start
        { 
            get { return air_gap_start;  }
        }

        public Point3D Air_Gap_End
        {
            get { return air_gap_end; }
        }

        public Snout(double snout_distance, double gantry_angle, double couch_rtn, Vector3D isocenter)
        {
            this.current_snout_distance = snout_distance;
            this.planned_snout_distance = snout_distance;
            this.gantry_angle = gantry_angle;
            this.couch_rotation = couch_rtn;
            this.isocenter = isocenter;
            Create3DGeometry();
        }

        private void Create3DGeometry()
        {
            //snout face center with gantry at 0            
            Vector3D snout_center_pos = new Vector3D(0, -planned_snout_distance, 0);

            //Calculating corners of snout cover

            //front/proximal face
            Vector3D bottomleftfront = Vector3D.Add(snout_center_pos, new Vector3D(snout_face_xmin, 0, snout_face_zmin));
            Vector3D bottomrightfront = Vector3D.Add(snout_center_pos, new Vector3D(snout_face_xmax, 0, snout_face_zmin));
            Vector3D topleftfront = Vector3D.Add(snout_center_pos, new Vector3D(snout_face_xmin, 0, snout_face_zmax));
            Vector3D toprightfront = Vector3D.Add(snout_center_pos, new Vector3D(snout_face_xmax, 0, snout_face_zmax));
            //back/distal end, assuming conical shape, size increases distally
            Vector3D bottomleftback = Vector3D.Add(snout_center_pos, new Vector3D(snout_end_xmin, -snout_depth, snout_end_zmin));
            Vector3D bottomrightback = Vector3D.Add(snout_center_pos, new Vector3D(snout_end_xmax, -snout_depth, snout_end_zmin));
            Vector3D topleftback = Vector3D.Add(snout_center_pos, new Vector3D(snout_end_xmin, -snout_depth, snout_end_zmax));
            Vector3D toprightback = Vector3D.Add(snout_center_pos, new Vector3D(snout_end_xmax, -snout_depth, snout_end_zmax));

            //gantry and couch rotation
            Matrix3D matrix3D = Matrix3D.Identity;
            Quaternion gantry_rot = new Quaternion(new Vector3D(0, 0, 1), gantry_angle);
            matrix3D.Rotate(gantry_rot);
            Quaternion couch_rot = new Quaternion(new Vector3D(0, 1 , 0), couch_rotation);
            matrix3D.Rotate(couch_rot);

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

            planned_snout_position = matrix3D.Transform(snout_center_pos);

            //Shifting by isocenter:
            planned_snout_position = Vector3D.Add(planned_snout_position, isocenter);

            bottomleftfront = Vector3D.Add(bottomleftfront, isocenter);
            bottomrightfront = Vector3D.Add(bottomrightfront, isocenter);
            topleftfront = Vector3D.Add(topleftfront, isocenter);
            toprightfront = Vector3D.Add(toprightfront, isocenter);
            bottomleftback = Vector3D.Add(bottomleftback, isocenter);
            bottomrightback = Vector3D.Add(bottomrightback, isocenter);
            topleftback = Vector3D.Add(topleftback, isocenter);
            toprightback = Vector3D.Add(toprightback, isocenter);

            //defining snout mesh for rotated snout
            snout_geometry = new Model3DGroup();
            GeometryModel3D snout_model = new GeometryModel3D();
            snout_mesh = new MeshGeometry3D();
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
            snout_geometry.Children.Add(snout_model);

            //creating frame
            Line3D front_bottom = new Line3D(Point3D.Add(Identity, bottomleftfront), Point3D.Add(Identity, bottomrightfront), Brushes.Black, 2);
            snout_geometry.Children.Add(front_bottom.GeometryModel3D);
            Line3D front_top = new Line3D(Point3D.Add(Identity, topleftfront), Point3D.Add(Identity, toprightfront), Brushes.Black, 2);
            snout_geometry.Children.Add(front_top.GeometryModel3D);
            Line3D front_right = new Line3D(Point3D.Add(Identity, bottomrightfront), Point3D.Add(Identity, toprightfront), Brushes.Black, 2);
            snout_geometry.Children.Add(front_right.GeometryModel3D);
            Line3D front_left = new Line3D(Point3D.Add(Identity, bottomleftfront), Point3D.Add(Identity, topleftfront), Brushes.Black, 2);
            snout_geometry.Children.Add(front_left.GeometryModel3D);
        }

        public void MoveTo(double new_snout_distance)
        {
            this.current_snout_distance = new_snout_distance;

            Vector3D snout_axis = Vector3D.Subtract(planned_snout_position, isocenter);

            snout_geometry.Transform = new TranslateTransform3D(Vector3D.Multiply(new_snout_distance / planned_snout_distance - 1, snout_axis));
        }

        public double CalculateAirGap(Structure body, double resolution_x, double resolution_y)
        {
            in_collision = false;
            air_gap_not_found = true;
            int x_counter = 0;
            air_gap = Double.MaxValue;
            VVector closest_collision_point = new VVector(0, 0, 0);
            VVector offset_at_snout = new VVector(0, 0, 0);

            //corner points of snout at planned snout distance
            Point3D snout_bottom_left = snout_mesh.Positions[0];
            Point3D snout_bottom_right = snout_mesh.Positions[1];
            Point3D snout_top_left = snout_mesh.Positions[2];

            //shift vectors to move along snout face
            Vector3D delta_x = Point3D.Subtract(snout_bottom_right, snout_bottom_left);
            Vector3D delta_y = Point3D.Subtract(snout_top_left, snout_bottom_left);

            //normalization to given resolution
            delta_x = Vector3D.Multiply(Vector3D.Divide(delta_x, delta_x.Length), resolution_x);
            delta_y = Vector3D.Multiply(Vector3D.Divide(delta_y, delta_y.Length), resolution_y);

            //vector perpenducular to snout face, its direction is towards patient
            Vector3D normal = Vector3D.CrossProduct(delta_y, delta_x);

            //normalization of snout face normal:
            Vector3D normal_unit = Vector3D.Divide(normal, normal.Length);
            normal = Vector3D.Multiply(normal_unit, current_snout_distance);
            Vector3D normal_planned = Vector3D.Multiply(normal_unit, planned_snout_distance);

            //bottom left corner point at current snout distance
            Point3D start_corner = Point3D.Add(snout_bottom_left,Vector3D.Subtract(normal_planned, normal));


            while (snout_face_xmin + x_counter * resolution_x <= snout_face_xmax)
            {
                int y_counter = 0;
                while (snout_face_zmin + y_counter * resolution_y <= snout_face_zmax)
                {
                    //start and end points for segment profile
                    Vector3D shift = Vector3D.Add(Vector3D.Multiply(delta_x, x_counter), Vector3D.Multiply(delta_y, y_counter));
                    Point3D start3D = Point3D.Add(start_corner, shift);
                    Point3D end3D = Point3D.Add(start3D, normal);

                    VVector start = new VVector(start3D.X, start3D.Y, start3D.Z);
                    VVector end = new VVector(end3D.X, end3D.Y, end3D.Z);

                    BitArray array = new BitArray(10 * (int)current_snout_distance);

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
                        air_gap_not_found = false;
                        
                        if (collision_index == 0)
                        {
                            in_collision = true;
                            break;
                        }
                        else
                        {
                            VVector collision_point = new VVector(profile[collision_index].Position.x, profile[collision_index].Position.y, profile[collision_index].Position.z);

                            Double current_air_gap_value = (collision_point - start).Length;

                            if (Math.Abs(current_air_gap_value) < Math.Abs(air_gap))
                            {
                                air_gap = current_air_gap_value;

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

            air_gap_start = new Point3D(closest_collision_point.x, closest_collision_point.y, closest_collision_point.z);
            air_gap_end = new Point3D(offset_at_snout.x, offset_at_snout.y, offset_at_snout.z);

            return air_gap;
        }
    }

    //GUI class
    public class WNDContent : UserControl
    {
        const double snout_pos_min = 0;
        const double snout_pos_max = 421;

        private Canvas canvas;
        private ComboBox field;
        private PerspectiveCamera camera;
        private Model3DGroup model3D;
        private System.Windows.Controls.Label snout_position;
        private System.Windows.Controls.Label view_angle;
        private TextBox txt_x;
        private TextBox txt_y;
        private Slider sl_snout_position;
        private System.Windows.Controls.Label air_gap;
        private Button btn_calculate;
        private Snout snout;

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

            System.Windows.Controls.Label lbl_patient = new System.Windows.Controls.Label();
            lbl_patient.Content = "Patient:";
            top_left_grid.Children.Add(lbl_patient);

            System.Windows.Controls.Label lbl_plan = new System.Windows.Controls.Label();
            lbl_plan.Content = "Plan:";
            lbl_plan.Margin = new Thickness(0, 15, 0, 0);
            top_left_grid.Children.Add(lbl_plan);

            System.Windows.Controls.Label patient = new System.Windows.Controls.Label();
            patient.Name = "patient";
            patient.Margin = new Thickness(50, 0, 0, 0);
            top_left_grid.Children.Add(patient);

            System.Windows.Controls.Label plan = new System.Windows.Controls.Label();
            plan.Name = "plan";
            plan.Margin = new Thickness(50, 15, 0, 0);
            top_left_grid.Children.Add(plan);

            System.Windows.Controls.Label lbl_field = new System.Windows.Controls.Label();
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

            System.Windows.Controls.Label lbl_airgap = new System.Windows.Controls.Label();
            lbl_airgap.Content = "Measure air gap at resolution:";
            lbl_airgap.Margin = new Thickness(0, 15, 0, 0);
            top_right_grid.Children.Add(lbl_airgap);

            System.Windows.Controls.Label lbl_x = new System.Windows.Controls.Label();
            lbl_x.Content = "x[mm]:";
            lbl_x.Margin = new Thickness(165, 0, 0, 0);
            top_right_grid.Children.Add(lbl_x);

            System.Windows.Controls.Label lbl_y = new System.Windows.Controls.Label();
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

            air_gap = new System.Windows.Controls.Label();
            air_gap.Content = "Calculated air gap =";
            air_gap.Margin = new Thickness(100, 40, 0, 0);
            top_right_grid.Children.Add(air_gap);

            top_right.Child = top_right_grid;

            //Canvas controls
            System.Windows.Controls.Label lbl_view_angle = new System.Windows.Controls.Label();
            lbl_view_angle.Content = "View angle[deg]: ";
            lbl_view_angle.Margin = new Thickness(5, 80, 0, 0);
            main_grid.Children.Add(lbl_view_angle);

            view_angle = new System.Windows.Controls.Label();
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

            System.Windows.Controls.Label lbl_snout_position = new System.Windows.Controls.Label();
            lbl_snout_position.Content = "Snout position[cm]: ";
            lbl_snout_position.Margin = new Thickness(5, 100, 0, 0);
            main_grid.Children.Add(lbl_snout_position);

            snout_position = new System.Windows.Controls.Label();
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

            // Create the TextBlock
            TextBlock myTextBlock = new TextBlock();

            // Set the row property of the TextBlock to 1
            Grid.SetRow(myTextBlock, 1);

            // Set the name and background of the TextBlock
            myTextBlock.Name = "Footer";
            myTextBlock.Background = new SolidColorBrush(Colors.PaleVioletRed);

            // Create the first Label with a Hyperlink
            Label label1 = new System.Windows.Controls.Label();
            Hyperlink hyperlink = new Hyperlink();
            hyperlink.NavigateUri = new Uri("http://medicalaffairs.varian.com/download/VarianLUSLA.pdf");
            hyperlink.RequestNavigate += Hyperlink_RequestNavigate;
            hyperlink.Inlines.Add("Bound by the terms of the Varian LUSLA");
            label1.Content = hyperlink;
            label1.Margin = new Thickness(0);

            // Add the Labels to the TextBlock
            myTextBlock.Inlines.Add(label1);

            main_grid.Children.Add(myTextBlock);
            myTextBlock.VerticalAlignment = VerticalAlignment.Bottom;
            myTextBlock.HorizontalAlignment = HorizontalAlignment.Stretch;

            this.Content = main_grid;
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            // for .NET Core you need to add UseShellExecute = true
            // see https://learn.microsoft.com/dotnet/api/system.diagnostics.processstartinfo.useshellexecute#property-value
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        private void Btn_help_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("This script visualizes a snout for proton machine. \r\n\r\n" +
                "It allows moving the snout and calculating air gap.\r\n\r\n" + 
                "Air gap is calculated using ray tracing from the snout down to the body. Raytracing can be done at a custom resolution.\r\n\r\n"+
                "Note: small value may require a few seconds to calculate depending on snout size");
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

            double air_gap_value = snout.CalculateAirGap(body, resolution_x, resolution_y);

            //Erase previously calculated airgap:
            if (model3D.Children.Count == 6) //includes calculated gap which is the last child
            {
                model3D.Children.RemoveAt(5);
                air_gap.Content = "Calculated air gap = ";
            }

            if (snout.Air_Gap_Not_Found)
            {
                //in collision
                air_gap.Content = "Calculated air gap = NOT FOUND";
            }
            else
            {
                if (!snout.In_Collision)
                {
                    Line3D line = new Line3D(snout.Air_Gap_Start, snout.Air_Gap_End, Brushes.Red, 3);
                    model3D.Children.Add(line.GeometryModel3D);

                    air_gap.Content = "Calculated air gap = " + air_gap_value.ToString("###.0") + " mm";
                }
                else
                {
                    air_gap.Content = "Calculated air gap = IN COLLISION";
                }
            }
 
        }

        private void ShowStartupMsg()
        {
            string EULA_TEXT = @"""
            VARIAN LIMITED USE SOFTWARE LICENSE AGREEMENT
            This Limited Use Software License Agreement (the ""Agreement"") is a legal agreement between you , the
            user (“You”), and Varian Medical Systems, Inc. (""Varian""). By downloading or otherwise accessing the
            software material, which includes source code (the ""Source Code"") and related software tools (collectively,
            the ""Software""), You are agreeing to be bound by the terms of this Agreement. If You are entering into this
            Agreement on behalf of an institution or company, You represent and warrant that You are authorized to do
            so. If You do not agree to the terms of this Agreement, You may not use the Software and must immediately
            destroy any Software You may have downloaded or copied.
            SOFTWARE LICENSE
            1. Grant of License. Varian grants to You a non-transferable, non-sublicensable license to use
            the Software solely as provided in Section 2 (Permitted Uses) below. Access to the Software will be
            facilitated through a source code repository provided by Varian.
            2. Permitted Uses. You may download, compile and use the Software, You may (but are not required to do
            so) suggest to Varian improvements or otherwise provide feedback to Varian with respect to the
            Software. You may modify the Software solely in support of such use, and You may upload such
            modified Software to Varian’s source code repository. Any derivation of the Software (including compiled
            binaries) must display prominently the terms and conditions of this Agreement in the interactive user
            interface, such that use of the Software cannot continue until the user has acknowledged having read
            this Agreement via click-through.
            3. Publications. Solely in connection with your use of the Software as permitted under this Agreement, You
            may make reference to this Software in connection with such use in academic research publications
            after notifying an authorized representative of Varian in writing in each instance. Notwithstanding the
            foregoing, You may not make reference to the Software in any way that may indicate or imply any
            approval or endorsement by Varian of the results of any use of the Software by You.
            4. Prohibited Uses. Under no circumstances are You permitted, allowed or authorized to distribute the
            Software or any modifications to the Software for any purpose, including, but not limited to, renting,
            selling, or leasing the Software or any modifications to the Software, for free or otherwise. You may not
            disclose the Software to any third party without the prior express written consent of an authorized
            representative of Varian. You may not reproduce, copy or disclose to others, in whole or in any part, the
            Software or modifications to the Software, except within Your own institution or company, as applicable,
            to facilitate Your permitted use of the Software. You agree that the Software will not be shipped,
            transferred or exported into any country in violation of the U.S. Export Administration Act (or any other
            law governing such matters) and that You will not utilize, in any other manner, the Software in
            violation of any applicable law.
            5. Intellectual Property Rights. All intellectual property rights in the Software and any modifications to the
            Software are owned solely and exclusively by Varian, and You shall have no ownership or other
            proprietary interest in the Software or any modifications. You hereby transfer and assign to Varian all
            right, title and interest in any such modifications to the Software that you may have made or contributed.
            You hereby waive any and all moral rights that you may have with respect to such modifications, and
            hereby waive any rights of attribution relating to any modifications of the Software. You acknowledge
            that Varian will have the sole right to commercialize and otherwise use, whether directly or through third
            parties, any modifications to the Software that you provide to Varian’s repository. Varian may make any
            use it determines to be appropriate with respect to any feedback, suggestions or other communications
            that You provide with respect to the Software or any modifications.
            6. No Support Obligations. Varian is under no obligation to provide any support or technical assistance in
            connection with the Software or any modifications. Any such support or technical assistance is entirely
            discretionary on the part of Varian, and may be discontinued at any time without liability.
            7. NO WARRANTIES. THE SOFTWARE AND ANY SUPPORT PROVIDED BY VARIAN ARE PROVIDED
            “AS IS” AND “WITH ALL FAULTS.” VARIAN DISCLAIMS ALL WARRANTIES, BOTH EXPRESS AND
            IMPLIED, INCLUDING BUT NOT LIMITED TO IMPLIED WARRANTIES OF MERCHANTABILITY,
            FITNESS FOR A PARTICULAR PURPOSE, AND NON-INFRINGEMENT WITH RESPECT TO THE
            SOFTWARE AND ANY SUPPORT. VARIAN DOES NOT WARRANT THAT THE OPERATION OF THE
            SOFTWARE WILL BE UNINTERRUPTED, ERROR FREE OR MEET YOUR SPECIFIC
            REQUIREMENTS OR INTENDED USE. THE AGENTS AND EMPLOYEES OF VARIAN ARE NOT
            AUTHORIZED TO MAKE MODIFICATIONS TO THIS PROVISION, OR PROVIDE ADDITIONAL
            WARRANTIES ON BEHALF OF VARIAN.
            8. No Regulatory Clearance. The Software is not cleared or approved for use by any regulatory body in any
            jurisdiction.
            9. Termination. You may terminate this Agreement, and the right to use the Software, at any time upon
            written notice to Varian. Varian may terminate this Agreement, and the right to use the Software, at any
            time upon notice to You in the event that Varian determines that you are not using the Software in
            accordance with this Agreement or have otherwise breached any provision of this Agreement. The
            Software, together with any modifications to it or any permitted archive copy thereof, shall be destroyed
            when no longer used in accordance with this Agreement, or when the right to use the Software is
            terminated.
            10. Limitation of Liability. IN NO EVENT SHALL VARIAN BE LIABLE FOR LOSS OF DATA, LOSS OF
            PROFITS, LOST SAVINGS, SPECIAL, INCIDENTAL, CONSEQUENTIAL, INDIRECT OR
            OTHER SIMILAR DAMAGES ARISING FROM BREACH OF WARRANTY, BREACH OF
            CONTRACT, NEGLIGENCE, OR OTHER LEGAL THEORY EVEN IF VARIAN OR ITS AGENT HAS
            BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES, OR FOR ANY CLAIM BY ANY OTHER
            PARTY.
            11. Indemnification. You will defend, indemnify and hold harmless Varian, its affiliates and their respective
            officers, directors, employees, sublicensees, contractors, users and agents from any and all claims,
            losses, liabilities, damages, expenses and costs (including attorneys’ fees and court costs) arising out of
            any third-party claims related to or arising from your use of the Software or any modifications to the
            Software.
            12. Assignment. You may not assign any of Your rights or obligations under this Agreement without the
            written consent of Varian.
            13. Governing Law. This Agreement will be governed and construed under the laws of the State of California
            and the United States of America without regard to conflicts of law provisions. The parties agree to the
            exclusive jurisdiction of the state and federal courts located in Santa Clara County, California with
            respect to any disputes under or relating to this Agreement.
            14. Entire Agreement. This Agreement is the entire agreement of the parties as to the subject matter and
            supersedes all prior written and oral agreements and understandings relating to same. The Agreement
            may only be modified or amended in a writing signed by the parties that makes specific reference to the
            Agreement and the provision the parties intend to modify or amend.""";

            var res = FlexibleMessageBox.Show(EULA_TEXT, "EULA Agreement", MessageBoxButtons.YesNo);
            if (res == DialogResult.No)
            {
                return;
            }
        }

        private void Field_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (model3D != null)
            {
                Initiate3DView();
            }
            //update controls
            air_gap.Content = "Calculated air gap = ";
            sl_snout_position.Value = (context.IonPlanSetup.Beams.ElementAt(field.SelectedIndex) as IonBeam).SnoutPosition * 10; //converting to mm
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

                snout.MoveTo(e.NewValue);
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
            //When creating 3D view canvas has only 1 child. After that it will have 2, second one  being the 3D view
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

            VVector v_isocenter = fields.ElementAt(field.SelectedIndex).IsocenterPosition;
            Vector3D isocenter = new Vector3D(v_isocenter.x, v_isocenter.y, v_isocenter.z);

            //adding patient body 3D in the model
            GeometryModel3D patientmodel = new GeometryModel3D();
            patientmodel.Geometry = body.MeshGeometry;
            DiffuseMaterial dm = new DiffuseMaterial();
            dm.Brush = new SolidColorBrush(Color.FromScRgb(1, 0.8f, 0.7176470588235294f, 0.4784313725490196f));
            patientmodel.Material = dm;
            model3D.Children.Add(patientmodel);

            //creating snout mesh at well know geometry when gantry at 0  --------------------------------------------

            //read required information from the plan, for a field selected in drop down
            double gantry_angle = fields.ElementAt(field.SelectedIndex).ControlPoints[0].GantryAngle;
            double plan_snout_distance = (fields.ElementAt(field.SelectedIndex) as IonBeam).SnoutPosition * 10; //converting to mm
            double couch_rtn = fields.ElementAt(field.SelectedIndex).ControlPoints[0].PatientSupportAngle;

            //apply couch rotation:

            //To apply couch rotation, we can rotate the BODY or the Snout geometry
            //For visualization, rotating BODY is probably most straightforward. It can be done by applying Transform3D to BODY Model3D
            //However, Transform3D does not rotate points of the Mesh3D, it only changes coordinate system for the Model3D
            //The BODY is also used for air gap calculation and for this calculation the Mesh 3D of the Snout and BODY geometries must be in the 
            //same coordinate system i.e. we would need to rotate all points of BODY Mesh3D. This is doable but a new Mesh3D must be created and there is no
            //way to make this Mesh3D part of Structure which is needed to call GetSegmentProfile for air gap calculation.
            //
            //Therefore, to apply couch rotation, the Snout geometry will be rotated, not BODY. The 3D view coordinate system is then connected with the couch.

            snout = new Snout(plan_snout_distance, gantry_angle, couch_rtn, isocenter);
            model3D.Children.Add(snout.Geometry);

            //CAX display
            Vector3D snout_parked = Vector3D.Multiply(Vector3D.Divide(Vector3D.Subtract(snout.Planned_Snout_Position,isocenter),snout.Snout_Distance),snout_pos_max);
            Line3D CAX = new Line3D(Point3D.Add(new Point3D(0, 0, 0), Vector3D.Add(snout_parked,isocenter)), new Point3D(isocenter.X, isocenter.Y, isocenter.Z), Brushes.YellowGreen, 1);
            model3D.Children.Add(CAX.GeometryModel3D);

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

            if ((Math.Abs(line.X) >= Math.Abs(line.Y)) && (Math.Abs(line.X) >= Math.Abs(line.Z)))
            {
                norm_to_line1 = new Vector3D( -(line.Y + line.Z)/line.X , 1, 1);
            }

            if ((Math.Abs(line.Y) >= Math.Abs(line.X)) && (Math.Abs(line.Y) >= Math.Abs(line.Z)))
            {
                norm_to_line1 = new Vector3D(1,-(line.X + line.Z) / line.Y, 1);
            }

            if ((Math.Abs(line.Z) >= Math.Abs(line.X)) && (Math.Abs(line.Z) >= Math.Abs(line.Y)))
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
        bool IsValidated = false;

        public Script()
        {
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context, System.Windows.Window window, ScriptEnvironment environment)
        {
            ShowStartupMsg();
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
            window.Title = "MAAS-ProtonSnoutCollision";

            if (!IsValidated)
            {
                window.Title += "* * * NOT VALIDATED FOR CLINICAL USE * * *";
            }

            


            

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

        private void ShowStartupMsg()
        {
            string EULA_TEXT = @"""
            VARIAN LIMITED USE SOFTWARE LICENSE AGREEMENT
            This Limited Use Software License Agreement (the ""Agreement"") is a legal agreement between you , the
            user (“You”), and Varian Medical Systems, Inc. (""Varian""). By downloading or otherwise accessing the
            software material, which includes source code (the ""Source Code"") and related software tools (collectively,
            the ""Software""), You are agreeing to be bound by the terms of this Agreement. If You are entering into this
            Agreement on behalf of an institution or company, You represent and warrant that You are authorized to do
            so. If You do not agree to the terms of this Agreement, You may not use the Software and must immediately
            destroy any Software You may have downloaded or copied.
            SOFTWARE LICENSE
            1. Grant of License. Varian grants to You a non-transferable, non-sublicensable license to use
            the Software solely as provided in Section 2 (Permitted Uses) below. Access to the Software will be
            facilitated through a source code repository provided by Varian.
            2. Permitted Uses. You may download, compile and use the Software, You may (but are not required to do
            so) suggest to Varian improvements or otherwise provide feedback to Varian with respect to the
            Software. You may modify the Software solely in support of such use, and You may upload such
            modified Software to Varian’s source code repository. Any derivation of the Software (including compiled
            binaries) must display prominently the terms and conditions of this Agreement in the interactive user
            interface, such that use of the Software cannot continue until the user has acknowledged having read
            this Agreement via click-through.
            3. Publications. Solely in connection with your use of the Software as permitted under this Agreement, You
            may make reference to this Software in connection with such use in academic research publications
            after notifying an authorized representative of Varian in writing in each instance. Notwithstanding the
            foregoing, You may not make reference to the Software in any way that may indicate or imply any
            approval or endorsement by Varian of the results of any use of the Software by You.
            4. Prohibited Uses. Under no circumstances are You permitted, allowed or authorized to distribute the
            Software or any modifications to the Software for any purpose, including, but not limited to, renting,
            selling, or leasing the Software or any modifications to the Software, for free or otherwise. You may not
            disclose the Software to any third party without the prior express written consent of an authorized
            representative of Varian. You may not reproduce, copy or disclose to others, in whole or in any part, the
            Software or modifications to the Software, except within Your own institution or company, as applicable,
            to facilitate Your permitted use of the Software. You agree that the Software will not be shipped,
            transferred or exported into any country in violation of the U.S. Export Administration Act (or any other
            law governing such matters) and that You will not utilize, in any other manner, the Software in
            violation of any applicable law.
            5. Intellectual Property Rights. All intellectual property rights in the Software and any modifications to the
            Software are owned solely and exclusively by Varian, and You shall have no ownership or other
            proprietary interest in the Software or any modifications. You hereby transfer and assign to Varian all
            right, title and interest in any such modifications to the Software that you may have made or contributed.
            You hereby waive any and all moral rights that you may have with respect to such modifications, and
            hereby waive any rights of attribution relating to any modifications of the Software. You acknowledge
            that Varian will have the sole right to commercialize and otherwise use, whether directly or through third
            parties, any modifications to the Software that you provide to Varian’s repository. Varian may make any
            use it determines to be appropriate with respect to any feedback, suggestions or other communications
            that You provide with respect to the Software or any modifications.
            6. No Support Obligations. Varian is under no obligation to provide any support or technical assistance in
            connection with the Software or any modifications. Any such support or technical assistance is entirely
            discretionary on the part of Varian, and may be discontinued at any time without liability.
            7. NO WARRANTIES. THE SOFTWARE AND ANY SUPPORT PROVIDED BY VARIAN ARE PROVIDED
            “AS IS” AND “WITH ALL FAULTS.” VARIAN DISCLAIMS ALL WARRANTIES, BOTH EXPRESS AND
            IMPLIED, INCLUDING BUT NOT LIMITED TO IMPLIED WARRANTIES OF MERCHANTABILITY,
            FITNESS FOR A PARTICULAR PURPOSE, AND NON-INFRINGEMENT WITH RESPECT TO THE
            SOFTWARE AND ANY SUPPORT. VARIAN DOES NOT WARRANT THAT THE OPERATION OF THE
            SOFTWARE WILL BE UNINTERRUPTED, ERROR FREE OR MEET YOUR SPECIFIC
            REQUIREMENTS OR INTENDED USE. THE AGENTS AND EMPLOYEES OF VARIAN ARE NOT
            AUTHORIZED TO MAKE MODIFICATIONS TO THIS PROVISION, OR PROVIDE ADDITIONAL
            WARRANTIES ON BEHALF OF VARIAN.
            8. No Regulatory Clearance. The Software is not cleared or approved for use by any regulatory body in any
            jurisdiction.
            9. Termination. You may terminate this Agreement, and the right to use the Software, at any time upon
            written notice to Varian. Varian may terminate this Agreement, and the right to use the Software, at any
            time upon notice to You in the event that Varian determines that you are not using the Software in
            accordance with this Agreement or have otherwise breached any provision of this Agreement. The
            Software, together with any modifications to it or any permitted archive copy thereof, shall be destroyed
            when no longer used in accordance with this Agreement, or when the right to use the Software is
            terminated.
            10. Limitation of Liability. IN NO EVENT SHALL VARIAN BE LIABLE FOR LOSS OF DATA, LOSS OF
            PROFITS, LOST SAVINGS, SPECIAL, INCIDENTAL, CONSEQUENTIAL, INDIRECT OR
            OTHER SIMILAR DAMAGES ARISING FROM BREACH OF WARRANTY, BREACH OF
            CONTRACT, NEGLIGENCE, OR OTHER LEGAL THEORY EVEN IF VARIAN OR ITS AGENT HAS
            BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES, OR FOR ANY CLAIM BY ANY OTHER
            PARTY.
            11. Indemnification. You will defend, indemnify and hold harmless Varian, its affiliates and their respective
            officers, directors, employees, sublicensees, contractors, users and agents from any and all claims,
            losses, liabilities, damages, expenses and costs (including attorneys’ fees and court costs) arising out of
            any third-party claims related to or arising from your use of the Software or any modifications to the
            Software.
            12. Assignment. You may not assign any of Your rights or obligations under this Agreement without the
            written consent of Varian.
            13. Governing Law. This Agreement will be governed and construed under the laws of the State of California
            and the United States of America without regard to conflicts of law provisions. The parties agree to the
            exclusive jurisdiction of the state and federal courts located in Santa Clara County, California with
            respect to any disputes under or relating to this Agreement.
            14. Entire Agreement. This Agreement is the entire agreement of the parties as to the subject matter and
            supersedes all prior written and oral agreements and understandings relating to same. The Agreement
            may only be modified or amended in a writing signed by the parties that makes specific reference to the
            Agreement and the provision the parties intend to modify or amend.""";

            var res = FlexibleMessageBox.Show(EULA_TEXT, "EULA Agreement", MessageBoxButtons.YesNo);
            if (res == DialogResult.No)
            {
                return;
            }
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

    public class FlexibleMessageBox
    {
        #region Public statics

        /// <summary>
        /// Defines the maximum width for all FlexibleMessageBox instances in percent of the working area.
        /// 
        /// Allowed values are 0.2 - 1.0 where: 
        /// 0.2 means:  The FlexibleMessageBox can be at most half as wide as the working area.
        /// 1.0 means:  The FlexibleMessageBox can be as wide as the working area.
        /// 
        /// Default is: 70% of the working area width.
        /// </summary>
        public static double MAX_WIDTH_FACTOR = 0.7;

        /// <summary>
        /// Defines the maximum height for all FlexibleMessageBox instances in percent of the working area.
        /// 
        /// Allowed values are 0.2 - 1.0 where: 
        /// 0.2 means:  The FlexibleMessageBox can be at most half as high as the working area.
        /// 1.0 means:  The FlexibleMessageBox can be as high as the working area.
        /// 
        /// Default is: 90% of the working area height.
        /// </summary>
        public static double MAX_HEIGHT_FACTOR = 0.9;

        /// <summary>
        /// Defines the font for all FlexibleMessageBox instances.
        /// 
        /// Default is: SystemFonts.MessageBoxFont
        /// </summary>
        public static Font FONT = SystemFonts.MessageBoxFont;

        #endregion

        #region Public show functions

        /// <summary>
        /// Shows the specified message box.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <returns>The dialog result.</returns>
        public static DialogResult Show(string text)
        {
            return FlexibleMessageBoxForm.Show(null, text, string.Empty, MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1);
        }

        /// <summary>
        /// Shows the specified message box.
        /// </summary>
        /// <param name="owner">The owner.</param>
        /// <param name="text">The text.</param>
        /// <returns>The dialog result.</returns>
        public static DialogResult Show(IWin32Window owner, string text)
        {
            return FlexibleMessageBoxForm.Show(owner, text, string.Empty, MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1);
        }

        /// <summary>
        /// Shows the specified message box.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <param name="caption">The caption.</param>
        /// <returns>The dialog result.</returns>
        public static DialogResult Show(string text, string caption)
        {
            return FlexibleMessageBoxForm.Show(null, text, caption, MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1);
        }

        /// <summary>
        /// Shows the specified message box.
        /// </summary>
        /// <param name="owner">The owner.</param>
        /// <param name="text">The text.</param>
        /// <param name="caption">The caption.</param>
        /// <returns>The dialog result.</returns>
        public static DialogResult Show(IWin32Window owner, string text, string caption)
        {
            return FlexibleMessageBoxForm.Show(owner, text, caption, MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1);
        }

        /// <summary>
        /// Shows the specified message box.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <param name="caption">The caption.</param>
        /// <param name="buttons">The buttons.</param>
        /// <returns>The dialog result.</returns>
        public static DialogResult Show(string text, string caption, MessageBoxButtons buttons)
        {
            return FlexibleMessageBoxForm.Show(null, text, caption, buttons, MessageBoxIcon.None, MessageBoxDefaultButton.Button1);
        }

        /// <summary>
        /// Shows the specified message box.
        /// </summary>
        /// <param name="owner">The owner.</param>
        /// <param name="text">The text.</param>
        /// <param name="caption">The caption.</param>
        /// <param name="buttons">The buttons.</param>
        /// <returns>The dialog result.</returns>
        public static DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons)
        {
            return FlexibleMessageBoxForm.Show(owner, text, caption, buttons, MessageBoxIcon.None, MessageBoxDefaultButton.Button1);
        }

        /// <summary>
        /// Shows the specified message box.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <param name="caption">The caption.</param>
        /// <param name="buttons">The buttons.</param>
        /// <param name="icon">The icon.</param>
        /// <returns></returns>
        public static DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            return FlexibleMessageBoxForm.Show(null, text, caption, buttons, icon, MessageBoxDefaultButton.Button1);
        }

        /// <summary>
        /// Shows the specified message box.
        /// </summary>
        /// <param name="owner">The owner.</param>
        /// <param name="text">The text.</param>
        /// <param name="caption">The caption.</param>
        /// <param name="buttons">The buttons.</param>
        /// <param name="icon">The icon.</param>
        /// <returns>The dialog result.</returns>
        public static DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            return FlexibleMessageBoxForm.Show(owner, text, caption, buttons, icon, MessageBoxDefaultButton.Button1);
        }

        /// <summary>
        /// Shows the specified message box.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <param name="caption">The caption.</param>
        /// <param name="buttons">The buttons.</param>
        /// <param name="icon">The icon.</param>
        /// <param name="defaultButton">The default button.</param>
        /// <returns>The dialog result.</returns>
        public static DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton)
        {
            return FlexibleMessageBoxForm.Show(null, text, caption, buttons, icon, defaultButton);
        }

        /// <summary>
        /// Shows the specified message box.
        /// </summary>
        /// <param name="owner">The owner.</param>
        /// <param name="text">The text.</param>
        /// <param name="caption">The caption.</param>
        /// <param name="buttons">The buttons.</param>
        /// <param name="icon">The icon.</param>
        /// <param name="defaultButton">The default button.</param>
        /// <returns>The dialog result.</returns>
        public static DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton)
        {
            return FlexibleMessageBoxForm.Show(owner, text, caption, buttons, icon, defaultButton);
        }

        #endregion

        #region Internal form class

        /// <summary>
        /// The form to show the customized message box.
        /// It is defined as an internal class to keep the public interface of the FlexibleMessageBox clean.
        /// </summary>
        class FlexibleMessageBoxForm : Form
        {
            #region Form-Designer generated code

            /// <summary>
            /// Erforderliche Designervariable.
            /// </summary>
            private System.ComponentModel.IContainer components = null;

            /// <summary>
            /// Verwendete Ressourcen bereinigen.
            /// </summary>
            /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
            protected override void Dispose(bool disposing)
            {
                if (disposing && (components != null))
                {
                    components.Dispose();
                }
                base.Dispose(disposing);
            }

            /// <summary>
            /// Erforderliche Methode für die Designerunterstützung.
            /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
            /// </summary>
            private void InitializeComponent()
            {
                this.components = new System.ComponentModel.Container();
                this.button1 = new System.Windows.Forms.Button();
                this.richTextBoxMessage = new System.Windows.Forms.RichTextBox();
                this.FlexibleMessageBoxFormBindingSource = new System.Windows.Forms.BindingSource(this.components);
                this.panel1 = new System.Windows.Forms.Panel();
                this.pictureBoxForIcon = new System.Windows.Forms.PictureBox();
                this.button2 = new System.Windows.Forms.Button();
                this.button3 = new System.Windows.Forms.Button();
                ((System.ComponentModel.ISupportInitialize)(this.FlexibleMessageBoxFormBindingSource)).BeginInit();
                this.panel1.SuspendLayout();
                ((System.ComponentModel.ISupportInitialize)(this.pictureBoxForIcon)).BeginInit();
                this.SuspendLayout();
                // 
                // button1
                // 
                this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
                this.button1.AutoSize = true;
                this.button1.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.button1.Location = new System.Drawing.Point(11, 67);
                this.button1.MinimumSize = new System.Drawing.Size(0, 24);
                this.button1.Name = "button1";
                this.button1.Size = new System.Drawing.Size(75, 24);
                this.button1.TabIndex = 2;
                this.button1.Text = "OK";
                this.button1.UseVisualStyleBackColor = true;
                this.button1.Visible = false;
                // 
                // richTextBoxMessage
                // 
                this.richTextBoxMessage.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                | System.Windows.Forms.AnchorStyles.Left)
                | System.Windows.Forms.AnchorStyles.Right)));
                this.richTextBoxMessage.BackColor = System.Drawing.Color.White;
                this.richTextBoxMessage.BorderStyle = System.Windows.Forms.BorderStyle.None;
                this.richTextBoxMessage.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.FlexibleMessageBoxFormBindingSource, "MessageText", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
                this.richTextBoxMessage.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                this.richTextBoxMessage.Location = new System.Drawing.Point(50, 26);
                this.richTextBoxMessage.Margin = new System.Windows.Forms.Padding(0);
                this.richTextBoxMessage.Name = "richTextBoxMessage";
                this.richTextBoxMessage.ReadOnly = true;
                this.richTextBoxMessage.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
                this.richTextBoxMessage.Size = new System.Drawing.Size(200, 20);
                this.richTextBoxMessage.TabIndex = 0;
                this.richTextBoxMessage.TabStop = false;
                this.richTextBoxMessage.Text = "<Message>";
                this.richTextBoxMessage.LinkClicked += new System.Windows.Forms.LinkClickedEventHandler(this.richTextBoxMessage_LinkClicked);
                // 
                // panel1
                // 
                this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                | System.Windows.Forms.AnchorStyles.Left)
                | System.Windows.Forms.AnchorStyles.Right)));
                this.panel1.BackColor = System.Drawing.Color.White;
                this.panel1.Controls.Add(this.pictureBoxForIcon);
                this.panel1.Controls.Add(this.richTextBoxMessage);
                this.panel1.Location = new System.Drawing.Point(-3, -4);
                this.panel1.Name = "panel1";
                this.panel1.Size = new System.Drawing.Size(268, 59);
                this.panel1.TabIndex = 1;
                // 
                // pictureBoxForIcon
                // 
                this.pictureBoxForIcon.BackColor = System.Drawing.Color.Transparent;
                this.pictureBoxForIcon.Location = new System.Drawing.Point(15, 19);
                this.pictureBoxForIcon.Name = "pictureBoxForIcon";
                this.pictureBoxForIcon.Size = new System.Drawing.Size(32, 32);
                this.pictureBoxForIcon.TabIndex = 8;
                this.pictureBoxForIcon.TabStop = false;
                // 
                // button2
                // 
                this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
                this.button2.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.button2.Location = new System.Drawing.Point(92, 67);
                this.button2.MinimumSize = new System.Drawing.Size(0, 24);
                this.button2.Name = "button2";
                this.button2.Size = new System.Drawing.Size(75, 24);
                this.button2.TabIndex = 3;
                this.button2.Text = "OK";
                this.button2.UseVisualStyleBackColor = true;
                this.button2.Visible = false;
                // 
                // button3
                // 
                this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
                this.button3.AutoSize = true;
                this.button3.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.button3.Location = new System.Drawing.Point(173, 67);
                this.button3.MinimumSize = new System.Drawing.Size(0, 24);
                this.button3.Name = "button3";
                this.button3.Size = new System.Drawing.Size(75, 24);
                this.button3.TabIndex = 0;
                this.button3.Text = "OK";
                this.button3.UseVisualStyleBackColor = true;
                this.button3.Visible = false;
                // 
                // FlexibleMessageBoxForm
                // 
                this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
                this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
                this.ClientSize = new System.Drawing.Size(260, 102);
                this.Controls.Add(this.button3);
                this.Controls.Add(this.button2);
                this.Controls.Add(this.panel1);
                this.Controls.Add(this.button1);
                this.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.FlexibleMessageBoxFormBindingSource, "CaptionText", true));
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.MinimumSize = new System.Drawing.Size(276, 140);
                this.Name = "FlexibleMessageBoxForm";
                this.ShowIcon = false;
                this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
                this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
                this.Text = "<Caption>";
                this.Shown += new System.EventHandler(this.FlexibleMessageBoxForm_Shown);
                ((System.ComponentModel.ISupportInitialize)(this.FlexibleMessageBoxFormBindingSource)).EndInit();
                this.panel1.ResumeLayout(false);
                ((System.ComponentModel.ISupportInitialize)(this.pictureBoxForIcon)).EndInit();
                this.ResumeLayout(false);
                this.PerformLayout();
            }

            private System.Windows.Forms.Button button1;
            private System.Windows.Forms.BindingSource FlexibleMessageBoxFormBindingSource;
            private System.Windows.Forms.RichTextBox richTextBoxMessage;
            private System.Windows.Forms.Panel panel1;
            private System.Windows.Forms.PictureBox pictureBoxForIcon;
            private System.Windows.Forms.Button button2;
            private System.Windows.Forms.Button button3;

            #endregion

            #region Private constants

            //These separators are used for the "copy to clipboard" standard operation, triggered by Ctrl + C (behavior and clipboard format is like in a standard MessageBox)
            private static readonly String STANDARD_MESSAGEBOX_SEPARATOR_LINES = "---------------------------\n";
            private static readonly String STANDARD_MESSAGEBOX_SEPARATOR_SPACES = "   ";

            //These are the possible buttons (in a standard MessageBox)
            private enum ButtonID { OK = 0, CANCEL, YES, NO, ABORT, RETRY, IGNORE };

            //These are the buttons texts for different languages. 
            //If you want to add a new language, add it here and in the GetButtonText-Function
            private enum TwoLetterISOLanguageID { en, de, es, it };
            private static readonly String[] BUTTON_TEXTS_ENGLISH_EN = { "OK", "Cancel", "&Yes", "&No", "&Abort", "&Retry", "&Ignore" }; //Note: This is also the fallback language
            private static readonly String[] BUTTON_TEXTS_GERMAN_DE = { "OK", "Abbrechen", "&Ja", "&Nein", "&Abbrechen", "&Wiederholen", "&Ignorieren" };
            private static readonly String[] BUTTON_TEXTS_SPANISH_ES = { "Aceptar", "Cancelar", "&Sí", "&No", "&Abortar", "&Reintentar", "&Ignorar" };
            private static readonly String[] BUTTON_TEXTS_ITALIAN_IT = { "OK", "Annulla", "&Sì", "&No", "&Interrompi", "&Riprova", "&Ignora" };

            #endregion

            #region Private members

            private MessageBoxDefaultButton defaultButton;
            private int visibleButtonsCount;
            private TwoLetterISOLanguageID languageID = TwoLetterISOLanguageID.en;

            #endregion

            #region Private constructor

            /// <summary>
            /// Initializes a new instance of the <see cref="FlexibleMessageBoxForm"/> class.
            /// </summary>
            private FlexibleMessageBoxForm()
            {
                InitializeComponent();

                //Try to evaluate the language. If this fails, the fallback language English will be used
                Enum.TryParse<TwoLetterISOLanguageID>(CultureInfo.CurrentUICulture.TwoLetterISOLanguageName, out this.languageID);

                this.KeyPreview = true;
                this.KeyUp += FlexibleMessageBoxForm_KeyUp;
            }

            #endregion

            #region Private helper functions

            /// <summary>
            /// Gets the string rows.
            /// </summary>
            /// <param name="message">The message.</param>
            /// <returns>The string rows as 1-dimensional array</returns>
            private static string[] GetStringRows(string message)
            {
                if (string.IsNullOrEmpty(message)) return null;

                var messageRows = message.Split(new char[] { '\n' }, StringSplitOptions.None);
                return messageRows;
            }

            /// <summary>
            /// Gets the button text for the CurrentUICulture language.
            /// Note: The fallback language is English
            /// </summary>
            /// <param name="buttonID">The ID of the button.</param>
            /// <returns>The button text</returns>
            private string GetButtonText(ButtonID buttonID)
            {
                var buttonTextArrayIndex = Convert.ToInt32(buttonID);

                switch (this.languageID)
                {
                    case TwoLetterISOLanguageID.de: return BUTTON_TEXTS_GERMAN_DE[buttonTextArrayIndex];
                    case TwoLetterISOLanguageID.es: return BUTTON_TEXTS_SPANISH_ES[buttonTextArrayIndex];
                    case TwoLetterISOLanguageID.it: return BUTTON_TEXTS_ITALIAN_IT[buttonTextArrayIndex];

                    default: return BUTTON_TEXTS_ENGLISH_EN[buttonTextArrayIndex];
                }
            }

            /// <summary>
            /// Ensure the given working area factor in the range of  0.2 - 1.0 where: 
            /// 
            /// 0.2 means:  20 percent of the working area height or width.
            /// 1.0 means:  100 percent of the working area height or width.
            /// </summary>
            /// <param name="workingAreaFactor">The given working area factor.</param>
            /// <returns>The corrected given working area factor.</returns>
            private static double GetCorrectedWorkingAreaFactor(double workingAreaFactor)
            {
                const double MIN_FACTOR = 0.2;
                const double MAX_FACTOR = 1.0;

                if (workingAreaFactor < MIN_FACTOR) return MIN_FACTOR;
                if (workingAreaFactor > MAX_FACTOR) return MAX_FACTOR;

                return workingAreaFactor;
            }

            /// <summary>
            /// Set the dialogs start position when given. 
            /// Otherwise center the dialog on the current screen.
            /// </summary>
            /// <param name="flexibleMessageBoxForm">The FlexibleMessageBox dialog.</param>
            /// <param name="owner">The owner.</param>
            private static void SetDialogStartPosition(FlexibleMessageBoxForm flexibleMessageBoxForm, IWin32Window owner)
            {
                //If no owner given: Center on current screen
                if (owner == null)
                {
                    var screen = Screen.FromPoint(Cursor.Position);
                    flexibleMessageBoxForm.StartPosition = FormStartPosition.Manual;
                    flexibleMessageBoxForm.Left = screen.Bounds.Left + screen.Bounds.Width / 2 - flexibleMessageBoxForm.Width / 2;
                    flexibleMessageBoxForm.Top = screen.Bounds.Top + screen.Bounds.Height / 2 - flexibleMessageBoxForm.Height / 2;
                }
            }

            /// <summary>
            /// Calculate the dialogs start size (Try to auto-size width to show longest text row).
            /// Also set the maximum dialog size. 
            /// </summary>
            /// <param name="flexibleMessageBoxForm">The FlexibleMessageBox dialog.</param>
            /// <param name="text">The text (the longest text row is used to calculate the dialog width).</param>
            /// <param name="text">The caption (this can also affect the dialog width).</param>
            private static void SetDialogSizes(FlexibleMessageBoxForm flexibleMessageBoxForm, string text, string caption)
            {
                //First set the bounds for the maximum dialog size
                flexibleMessageBoxForm.MaximumSize = new Size(Convert.ToInt32(SystemInformation.WorkingArea.Width * FlexibleMessageBoxForm.GetCorrectedWorkingAreaFactor(MAX_WIDTH_FACTOR)),
                                                              Convert.ToInt32(SystemInformation.WorkingArea.Height * FlexibleMessageBoxForm.GetCorrectedWorkingAreaFactor(MAX_HEIGHT_FACTOR)));

                //Get rows. Exit if there are no rows to render...
                var stringRows = GetStringRows(text);
                if (stringRows == null) return;

                //Calculate whole text height
                var textHeight = TextRenderer.MeasureText(text, FONT).Height;

                //Calculate width for longest text line
                const int SCROLLBAR_WIDTH_OFFSET = 15;
                var longestTextRowWidth = stringRows.Max(textForRow => TextRenderer.MeasureText(textForRow, FONT).Width);
                var captionWidth = TextRenderer.MeasureText(caption, SystemFonts.CaptionFont).Width;
                var textWidth = Math.Max(longestTextRowWidth + SCROLLBAR_WIDTH_OFFSET, captionWidth);

                //Calculate margins
                var marginWidth = flexibleMessageBoxForm.Width - flexibleMessageBoxForm.richTextBoxMessage.Width;
                var marginHeight = flexibleMessageBoxForm.Height - flexibleMessageBoxForm.richTextBoxMessage.Height;

                //Set calculated dialog size (if the calculated values exceed the maximums, they were cut by windows forms automatically)
                flexibleMessageBoxForm.Size = new Size(textWidth + marginWidth,
                                                       textHeight + marginHeight);
            }

            /// <summary>
            /// Set the dialogs icon. 
            /// When no icon is used: Correct placement and width of rich text box.
            /// </summary>
            /// <param name="flexibleMessageBoxForm">The FlexibleMessageBox dialog.</param>
            /// <param name="icon">The MessageBoxIcon.</param>
            private static void SetDialogIcon(FlexibleMessageBoxForm flexibleMessageBoxForm, MessageBoxIcon icon)
            {
                switch (icon)
                {
                    case MessageBoxIcon.Information:
                        flexibleMessageBoxForm.pictureBoxForIcon.Image = SystemIcons.Information.ToBitmap();
                        break;
                    case MessageBoxIcon.Warning:
                        flexibleMessageBoxForm.pictureBoxForIcon.Image = SystemIcons.Warning.ToBitmap();
                        break;
                    case MessageBoxIcon.Error:
                        flexibleMessageBoxForm.pictureBoxForIcon.Image = SystemIcons.Error.ToBitmap();
                        break;
                    case MessageBoxIcon.Question:
                        flexibleMessageBoxForm.pictureBoxForIcon.Image = SystemIcons.Question.ToBitmap();
                        break;
                    default:
                        //When no icon is used: Correct placement and width of rich text box.
                        flexibleMessageBoxForm.pictureBoxForIcon.Visible = false;
                        flexibleMessageBoxForm.richTextBoxMessage.Left -= flexibleMessageBoxForm.pictureBoxForIcon.Width;
                        flexibleMessageBoxForm.richTextBoxMessage.Width += flexibleMessageBoxForm.pictureBoxForIcon.Width;
                        break;
                }
            }

            /// <summary>
            /// Set dialog buttons visibilities and texts. 
            /// Also set a default button.
            /// </summary>
            /// <param name="flexibleMessageBoxForm">The FlexibleMessageBox dialog.</param>
            /// <param name="buttons">The buttons.</param>
            /// <param name="defaultButton">The default button.</param>
            private static void SetDialogButtons(FlexibleMessageBoxForm flexibleMessageBoxForm, MessageBoxButtons buttons, MessageBoxDefaultButton defaultButton)
            {
                //Set the buttons visibilities and texts
                switch (buttons)
                {
                    case MessageBoxButtons.AbortRetryIgnore:
                        flexibleMessageBoxForm.visibleButtonsCount = 3;

                        flexibleMessageBoxForm.button1.Visible = true;
                        flexibleMessageBoxForm.button1.Text = flexibleMessageBoxForm.GetButtonText(ButtonID.ABORT);
                        flexibleMessageBoxForm.button1.DialogResult = DialogResult.Abort;

                        flexibleMessageBoxForm.button2.Visible = true;
                        flexibleMessageBoxForm.button2.Text = flexibleMessageBoxForm.GetButtonText(ButtonID.RETRY);
                        flexibleMessageBoxForm.button2.DialogResult = DialogResult.Retry;

                        flexibleMessageBoxForm.button3.Visible = true;
                        flexibleMessageBoxForm.button3.Text = flexibleMessageBoxForm.GetButtonText(ButtonID.IGNORE);
                        flexibleMessageBoxForm.button3.DialogResult = DialogResult.Ignore;

                        flexibleMessageBoxForm.ControlBox = false;
                        break;

                    case MessageBoxButtons.OKCancel:
                        flexibleMessageBoxForm.visibleButtonsCount = 2;

                        flexibleMessageBoxForm.button2.Visible = true;
                        flexibleMessageBoxForm.button2.Text = flexibleMessageBoxForm.GetButtonText(ButtonID.OK);
                        flexibleMessageBoxForm.button2.DialogResult = DialogResult.OK;

                        flexibleMessageBoxForm.button3.Visible = true;
                        flexibleMessageBoxForm.button3.Text = flexibleMessageBoxForm.GetButtonText(ButtonID.CANCEL);
                        flexibleMessageBoxForm.button3.DialogResult = DialogResult.Cancel;

                        flexibleMessageBoxForm.CancelButton = flexibleMessageBoxForm.button3;
                        break;

                    case MessageBoxButtons.RetryCancel:
                        flexibleMessageBoxForm.visibleButtonsCount = 2;

                        flexibleMessageBoxForm.button2.Visible = true;
                        flexibleMessageBoxForm.button2.Text = flexibleMessageBoxForm.GetButtonText(ButtonID.RETRY);
                        flexibleMessageBoxForm.button2.DialogResult = DialogResult.Retry;

                        flexibleMessageBoxForm.button3.Visible = true;
                        flexibleMessageBoxForm.button3.Text = flexibleMessageBoxForm.GetButtonText(ButtonID.CANCEL);
                        flexibleMessageBoxForm.button3.DialogResult = DialogResult.Cancel;

                        flexibleMessageBoxForm.CancelButton = flexibleMessageBoxForm.button3;
                        break;

                    case MessageBoxButtons.YesNo:
                        flexibleMessageBoxForm.visibleButtonsCount = 2;

                        flexibleMessageBoxForm.button2.Visible = true;
                        flexibleMessageBoxForm.button2.Text = flexibleMessageBoxForm.GetButtonText(ButtonID.YES);
                        flexibleMessageBoxForm.button2.DialogResult = DialogResult.Yes;

                        flexibleMessageBoxForm.button3.Visible = true;
                        flexibleMessageBoxForm.button3.Text = flexibleMessageBoxForm.GetButtonText(ButtonID.NO);
                        flexibleMessageBoxForm.button3.DialogResult = DialogResult.No;

                        flexibleMessageBoxForm.ControlBox = false;
                        break;

                    case MessageBoxButtons.YesNoCancel:
                        flexibleMessageBoxForm.visibleButtonsCount = 3;

                        flexibleMessageBoxForm.button1.Visible = true;
                        flexibleMessageBoxForm.button1.Text = flexibleMessageBoxForm.GetButtonText(ButtonID.YES);
                        flexibleMessageBoxForm.button1.DialogResult = DialogResult.Yes;

                        flexibleMessageBoxForm.button2.Visible = true;
                        flexibleMessageBoxForm.button2.Text = flexibleMessageBoxForm.GetButtonText(ButtonID.NO);
                        flexibleMessageBoxForm.button2.DialogResult = DialogResult.No;

                        flexibleMessageBoxForm.button3.Visible = true;
                        flexibleMessageBoxForm.button3.Text = flexibleMessageBoxForm.GetButtonText(ButtonID.CANCEL);
                        flexibleMessageBoxForm.button3.DialogResult = DialogResult.Cancel;

                        flexibleMessageBoxForm.CancelButton = flexibleMessageBoxForm.button3;
                        break;

                    case MessageBoxButtons.OK:
                    default:
                        flexibleMessageBoxForm.visibleButtonsCount = 1;
                        flexibleMessageBoxForm.button3.Visible = true;
                        flexibleMessageBoxForm.button3.Text = flexibleMessageBoxForm.GetButtonText(ButtonID.OK);
                        flexibleMessageBoxForm.button3.DialogResult = DialogResult.OK;

                        flexibleMessageBoxForm.CancelButton = flexibleMessageBoxForm.button3;
                        break;
                }

                //Set default button (used in FlexibleMessageBoxForm_Shown)
                flexibleMessageBoxForm.defaultButton = defaultButton;
            }

            #endregion

            #region Private event handlers

            /// <summary>
            /// Handles the Shown event of the FlexibleMessageBoxForm control.
            /// </summary>
            /// <param name="sender">The source of the event.</param>
            /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
            private void FlexibleMessageBoxForm_Shown(object sender, EventArgs e)
            {
                int buttonIndexToFocus = 1;
                Button buttonToFocus;

                //Set the default button...
                switch (this.defaultButton)
                {
                    case MessageBoxDefaultButton.Button1:
                    default:
                        buttonIndexToFocus = 1;
                        break;
                    case MessageBoxDefaultButton.Button2:
                        buttonIndexToFocus = 2;
                        break;
                    case MessageBoxDefaultButton.Button3:
                        buttonIndexToFocus = 3;
                        break;
                }

                if (buttonIndexToFocus > this.visibleButtonsCount) buttonIndexToFocus = this.visibleButtonsCount;

                if (buttonIndexToFocus == 3)
                {
                    buttonToFocus = this.button3;
                }
                else if (buttonIndexToFocus == 2)
                {
                    buttonToFocus = this.button2;
                }
                else
                {
                    buttonToFocus = this.button1;
                }

                buttonToFocus.Focus();
            }

            /// <summary>
            /// Handles the LinkClicked event of the richTextBoxMessage control.
            /// </summary>
            /// <param name="sender">The source of the event.</param>
            /// <param name="e">The <see cref="System.Windows.Forms.LinkClickedEventArgs"/> instance containing the event data.</param>
            private void richTextBoxMessage_LinkClicked(object sender, LinkClickedEventArgs e)
            {
                try
                {
                    Cursor.Current = Cursors.WaitCursor;
                    Process.Start(e.LinkText);
                }
                catch (Exception)
                {
                    //Let the caller of FlexibleMessageBoxForm decide what to do with this exception...
                    throw;
                }
                finally
                {
                    Cursor.Current = Cursors.Default;
                }

            }

            /// <summary>
            /// Handles the KeyUp event of the richTextBoxMessage control.
            /// </summary>
            /// <param name="sender">The source of the event.</param>
            /// <param name="e">The <see cref="System.Windows.Forms.KeyEventArgs"/> instance containing the event data.</param>
            void FlexibleMessageBoxForm_KeyUp(object sender, KeyEventArgs e)
            {
                //Handle standard key strikes for clipboard copy: "Ctrl + C" and "Ctrl + Insert"
                if (e.Control && (e.KeyCode == Keys.C || e.KeyCode == Keys.Insert))
                {
                    var buttonsTextLine = (this.button1.Visible ? this.button1.Text + STANDARD_MESSAGEBOX_SEPARATOR_SPACES : string.Empty)
                                        + (this.button2.Visible ? this.button2.Text + STANDARD_MESSAGEBOX_SEPARATOR_SPACES : string.Empty)
                                        + (this.button3.Visible ? this.button3.Text + STANDARD_MESSAGEBOX_SEPARATOR_SPACES : string.Empty);

                    //Build same clipboard text like the standard .Net MessageBox
                    var textForClipboard = STANDARD_MESSAGEBOX_SEPARATOR_LINES
                                         + this.Text + Environment.NewLine
                                         + STANDARD_MESSAGEBOX_SEPARATOR_LINES
                                         + this.richTextBoxMessage.Text + Environment.NewLine
                                         + STANDARD_MESSAGEBOX_SEPARATOR_LINES
                                         + buttonsTextLine.Replace("&", string.Empty) + Environment.NewLine
                                         + STANDARD_MESSAGEBOX_SEPARATOR_LINES;

                    //Set text in clipboard
                    Clipboard.SetText(textForClipboard);
                }
            }

            #endregion

            #region Properties (only used for binding)

            /// <summary>
            /// The text that is been used for the heading.
            /// </summary>
            public string CaptionText { get; set; }

            /// <summary>
            /// The text that is been used in the FlexibleMessageBoxForm.
            /// </summary>
            public string MessageText { get; set; }

            #endregion

            #region Public show function

            /// <summary>
            /// Shows the specified message box.
            /// </summary>
            /// <param name="owner">The owner.</param>
            /// <param name="text">The text.</param>
            /// <param name="caption">The caption.</param>
            /// <param name="buttons">The buttons.</param>
            /// <param name="icon">The icon.</param>
            /// <param name="defaultButton">The default button.</param>
            /// <returns>The dialog result.</returns>
            public static DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton)
            {
                //Create a new instance of the FlexibleMessageBox form
                var flexibleMessageBoxForm = new FlexibleMessageBoxForm();
                flexibleMessageBoxForm.ShowInTaskbar = false;

                //Bind the caption and the message text
                flexibleMessageBoxForm.CaptionText = caption;
                flexibleMessageBoxForm.MessageText = text;
                flexibleMessageBoxForm.FlexibleMessageBoxFormBindingSource.DataSource = flexibleMessageBoxForm;

                //Set the buttons visibilities and texts. Also set a default button.
                SetDialogButtons(flexibleMessageBoxForm, buttons, defaultButton);

                //Set the dialogs icon. When no icon is used: Correct placement and width of rich text box.
                SetDialogIcon(flexibleMessageBoxForm, icon);

                //Set the font for all controls
                flexibleMessageBoxForm.Font = FONT;
                flexibleMessageBoxForm.richTextBoxMessage.Font = FONT;

                //Calculate the dialogs start size (Try to auto-size width to show longest text row). Also set the maximum dialog size. 
                SetDialogSizes(flexibleMessageBoxForm, text, caption);

                //Set the dialogs start position when given. Otherwise center the dialog on the current screen.
                SetDialogStartPosition(flexibleMessageBoxForm, owner);

                //Show the dialog
                return flexibleMessageBoxForm.ShowDialog(owner);
            }

            #endregion
        } //class FlexibleMessageBoxForm
    
}
