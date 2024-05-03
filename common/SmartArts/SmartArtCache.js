/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

(function (window)
{
	const c_oAscSmartArtTypesToNameBinRelationShip = {};
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.AccentedPicture] = "Accented_Picture";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.AccentProcess] = "Accent_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.AlternatingFlow] = "Alternating_Flow";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.AlternatingHexagonList] = "Alternating_Hexagons";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.AlternatingPictureBlocks] = "Alternating_Picture_Blocks";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.AlternatingPictureCircles] = "Alternating_Picture_Circles";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.ArchitectureLayout] = "Architecture_Layout";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.ArrowRibbon] = "Arrow_Ribbon";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.AscendingPictureAccentProcess] = "Ascending_Picture_Accent_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.Balance] = "Balance";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BasicBendingProcess] = "Basic_Bending_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BasicBlockList] = "Basic_Block_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BasicChevronProcess] = "Basic_Chevron_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BasicCycle] = "Basic_Cycle";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BasicMatrix] = "Basic_Matrix";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BasicPie] = "Basic_Pie";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BasicProcess] = "Basic_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BasicPyramid] = "Basic_Pyramid";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BasicRadial] = "Basic_Radial";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BasicTarget] = "Basic_Target";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BasicTimeline] = "Basic_Timeline";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BasicVenn] = "Basic_Venn";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BendingPictureAccentList] = "Bending_Picture_Accent_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BendingPictureBlocks] = "Bending_Picture_Blocks";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BendingPictureCaption] = "Bending_Picture_Caption";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BendingPictureCaptionList] = "Bending_Picture_Caption_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BendingPictureSemiTransparentText] = "Bending_Picture_Semi-Transparent_Text";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BlockCycle] = "Block_Cycle";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.BubblePictureList] = "Bubble_Picture_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.CaptionedPictures] = "Captioned_Pictures";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.ChevronAccentProcess] = "Chevron_Accent_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.ChevronList] = "Chevron_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.CircleAccentTimeline] = "Circle_Accent_Timeline";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.CircleArrowProcess] = "Circle_Arrow_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.CirclePictureHierarchy] = "Circle_Picture_Hierarchy";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.CircleProcess] = "Circle_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.CircleRelationship] = "Circle_Relationship";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.CircularBendingProcess] = "Circular_Bending_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.CircularPictureCallout] = "Circular_Picture_Callout";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.ClosedChevronProcess] = "Closed_Chevron_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.ContinuousArrowProcess] = "Continuous_Arrow_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.ContinuousBlockProcess] = "Continuous_Block_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.ContinuousCycle] = "Continuous_Cycle";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.ContinuousPictureList] = "Continuous_Picture_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.ConvergingArrows] = "Converging_Arrows";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.ConvergingRadial] = "Converging_Radial";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.ConvergingText] = "Converging_Text";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.CounterbalanceArrows] = "Counterbalance_Arrows";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.CycleMatrix] = "Cycle_Matrix";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.DescendingBlockList] = "Descending_Block_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.DescendingProcess] = "Descending_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.DetailedProcess] = "Detailed_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.DivergingArrows] = "Diverging_Arrows";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.DivergingRadial] = "Diverging_Radial_";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.Equation] = "Equation";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.FramedTextPicture] = "Framed_Text_Picture";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.Funnel] = "Funnel";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.Gear] = "Gear";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.GridMatrix] = "Grid_Matrix";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.GroupedList] = "Grouped_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.HalfCircleOrganizationChart] = "Half_Circle_Organization_Chart";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.HexagonCluster] = "Hexagon_Cluster";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.HexagonRadial] = "Hexagon_Radial";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.Hierarchy] = "Hierarchy";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.HierarchyList] = "Hierarchy_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.HorizontalBulletList] = "Horizontal_Bullet_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.HorizontalHierarchy] = "Horizontal_Hierarchy";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.HorizontalLabeledHierarchy] = "Horizontal_Labeled_Hierarchy";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.HorizontalMultiLevelHierarchy] = "Horizontal_Multi-Level_Hierarchy";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.HorizontalOrganizationChart] = "Horizontal_Organization_Chart";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.HorizontalPictureList] = "Horizontal_Picture_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.IncreasingArrowsProcess] = "Increasing_Arrows_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.IncreasingCircleProcess] = "Increasing_Circle_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.InterconnectedBlockProcess] = "Interconnected_Block_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.InterconnectedRings] = "Interconnected_Rings";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.InvertedPyramid] = "Inverted_Pyramid";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.LabeledHierarchy] = "Labeled_Hierarchy";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.LinearVenn] = "Linear_Venn";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.LinedList] = "Lined_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.MultiDirectionalCycle] = "Multidirectional_Cycle";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.NameAndTitleOrganizationChart] = "Name_and_Title_Organization_Chart";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.NestedTarget] = "Nested_Target";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.NonDirectionalCycle] = "Nondirectional_Cycle";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.OpposingArrows] = "Opposing_Arrows";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.OpposingIdeas] = "Opposing_Ideas";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.OrganizationChart] = "Organization_Chart";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.PhasedProcess] = "Phased_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.PictureAccentBlocks] = "Picture_Accent_Blocks";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.PictureAccentList] = "Picture_Accent_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.PictureAccentProcess] = "Picture_Accent_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.PictureCaptionList] = "Picture_Caption_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.PictureFrame] = "Picture_Frame";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.PictureGrid] = "Picture_Grid";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.PictureLineup] = "Picture_Lineup";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.PictureOrganizationChart] = "Picture_Organization_Chart";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.PictureStrips] = "Picture_Strips";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.PieProcess] = "Pie_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.PlusAndMinus] = "Plus_and_Minus";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.ProcessArrows] = "Process_Arrows";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.ProcessList] = "Process_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.PyramidList] = "Pyramid_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.RadialCluster] = "Radial_Cluster";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.RadialCycle] = "Radial_Cycle";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.RadialList] = "Radial_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.RadialPictureList] = "Radial_Picture_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.RadialVenn] = "Radial_Venn";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.RandomToResultProcess] = "Random_to_Result_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.RepeatingBendingProcess] = "Repeating_Bending_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.ReverseList] = "Reverse_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.SegmentedCycle] = "Segmented_Cycle";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.SegmentedProcess] = "Segmented_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.SegmentedPyramid] = "Segmented_Pyramid";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.SnapshotPictureList] = "Snapshot_Picture_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.SpiralPicture] = "Spiral_Picture";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.SquareAccentList] = "Square_Accent_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.StackedList] = "Stacked_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.StackedVenn] = "Stacked_Venn";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.StaggeredProcess] = "Staggered_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.StepDownProcess] = "Step_Down_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.StepUpProcess] = "Step_Up_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.SubStepProcess] = "Sub-Step_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.TabbedArc] = "Tabbed_Arc";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.TableHierarchy] = "Table_Hierarchy";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.TableList] = "Table_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.TabList] = "Tab_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.TargetList] = "Target_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.TextCycle] = "Text_Cycle";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.ThemePictureAccent] = "Theme_Picture_Accent";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.ThemePictureAlternatingAccent] = "Theme_Picture_Alternating_Accent";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.ThemePictureGrid] = "Theme_Picture_Grid";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.TitledMatrix] = "Titled_Matrix";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.TitledPictureAccentList] = "Titled_Picture_Accent_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.TitledPictureBlocks] = "Titled_Picture_Blocks";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.TitlePictureLineup] = "Title_Picture_Lineup";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.TrapezoidList] = "Trapezoid_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.UpwardArrow] = "Upward_Arrow";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.VaryingWidthList] = "Varying_Width_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.VerticalAccentList] = "Vertical_Accent_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.VerticalArrowList] = "Vertical_Arrow_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.VerticalBendingProcess] = "Vertical_Bending_Process";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.VerticalBlockList] = "Vertical_Block_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.VerticalBoxList] = "Vertical_Box_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.VerticalBracketList] = "Vertical_Bracket_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.VerticalBulletList] = "Vertical_Bullet_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.VerticalChevronList] = "Vertical_Chevron_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.VerticalCircleList] = "Vertical_Circle_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.VerticalCurvedList] = "Vertical_Curved_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.VerticalEquation] = "Vertical_Equation";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.VerticalPictureAccentList] = "Vertical_Picture_Accent_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.VerticalPictureList] = "Vertical_Picture_List";
	c_oAscSmartArtTypesToNameBinRelationShip[Asc.c_oAscSmartArtTypes.VerticalProcess] = "Vertical_Process";

	function CSmartArtBinCache()
	{
		this.dataBin = {};
		this.drawingBin = {
			shifts: {},
			bin   : undefined
		};
	}

	CSmartArtBinCache.prototype.getDataBinary = function (nSmartArtType)
	{
		return this.dataBin[nSmartArtType];
	};
	CSmartArtBinCache.prototype.getDrawingInfo = function (nSmartArtType)
	{
		const nShift = this.drawingBin.shifts[nSmartArtType];
		if (nShift !== undefined)
		{
			return {pos: nShift, bin: this.drawingBin.bin};
		}
	};

	CSmartArtBinCache.prototype.checkLoadDrawing = function ()
	{
		const oThis = this;
		return new Promise(function (resolve)
		{
			if (oThis.drawingBin.bin)
			{
				resolve();
			}
			else if (!(window["NATIVE_EDITOR_ENJINE"] || oThis.drawingBin.bin === null))
			{
				oThis.drawingBin.bin = null;
				AscCommon.loadFileContent('../../../../sdkjs/common/SmartArts/SmartArtDrawing/SmartArtDrawings.bin', function (httpRequest)
				{
					if (httpRequest && httpRequest.response)
					{
						const arrStream = AscCommon.initStreamFromResponse(httpRequest);
						AscCommon.g_oBinarySmartArts.initDrawingFromBin(arrStream);
						resolve();
					}
				}, 'arraybuffer');
			}
		});
	};

	CSmartArtBinCache.prototype.initDrawingFromBin = function (arrStream)
	{
		const oFileStream = new AscCommon.FileStream(arrStream, arrStream.length);

		oFileStream.GetUChar();
		let nLength = oFileStream.GetULong();
		this.drawingBin.bin = new Uint8Array(oFileStream.GetBuffer(nLength));

		oFileStream.GetUChar();
		const nEnd = oFileStream.cur + oFileStream.GetULong() + 4;
		while (nEnd > oFileStream.cur)
		{
			const nType = oFileStream.GetUChar();
			this.drawingBin.shifts[nType] = oFileStream.GetULong();
		}
	};

	CSmartArtBinCache.prototype.initDataFromBin = function (nSmartArtType, arrStream)
	{
		this.dataBin[nSmartArtType] = new Uint8Array(arrStream);
	};

	CSmartArtBinCache.prototype.checkLoadData = function (nSmartArtType)
	{
		const oThis = this;
		return new Promise(function (resolve)
		{
			if (oThis.dataBin[nSmartArtType])
			{
				resolve();
			}
			else if (!(window["NATIVE_EDITOR_ENJINE"] || oThis.dataBin[nSmartArtType] === null))
			{
				oThis.dataBin[nSmartArtType] = null;
				const sFileName = c_oAscSmartArtTypesToNameBinRelationShip[nSmartArtType];
				AscCommon.loadFileContent('../../../../sdkjs/common/SmartArts/SmartArtData/' + sFileName + '.bin', function (httpRequest)
				{
					if (httpRequest && httpRequest.response)
					{
						const arrStream = AscCommon.initStreamFromResponse(httpRequest);
						AscCommon.g_oBinarySmartArts.initDataFromBin(nSmartArtType, arrStream);
						resolve();
					}
				}, 'arraybuffer');
			}
		});
	};

	window["AscCommon"] = window.AscCommon = window["AscCommon"] || {};

	window["AscCommon"].g_oBinarySmartArts = new CSmartArtBinCache();
}(window));
