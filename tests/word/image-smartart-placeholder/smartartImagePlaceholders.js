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
QUnit.config.autostart = false;
$(function () {

	const arrSmartArtTypesWithoutImagePlaceholders = [
		Asc.c_oAscSmartArtTypes.Balance,
		Asc.c_oAscSmartArtTypes.BlockCycle,
		Asc.c_oAscSmartArtTypes.StackedVenn,
		Asc.c_oAscSmartArtTypes.VerticalEquation,
		Asc.c_oAscSmartArtTypes.VerticalBlockList,
		Asc.c_oAscSmartArtTypes.VerticalBendingProcess,
		Asc.c_oAscSmartArtTypes.VerticalBulletList,
		Asc.c_oAscSmartArtTypes.VerticalCurvedList,
		Asc.c_oAscSmartArtTypes.VerticalProcess,
		Asc.c_oAscSmartArtTypes.VerticalBoxList,
		Asc.c_oAscSmartArtTypes.VerticalCircleList,
		Asc.c_oAscSmartArtTypes.VerticalArrowList,
		Asc.c_oAscSmartArtTypes.VerticalChevronList,
		Asc.c_oAscSmartArtTypes.VerticalAccentList,
		Asc.c_oAscSmartArtTypes.NestedTarget,
		Asc.c_oAscSmartArtTypes.Funnel,
		Asc.c_oAscSmartArtTypes.UpwardArrow,
		Asc.c_oAscSmartArtTypes.IncreasingArrowsProcess,
		Asc.c_oAscSmartArtTypes.StepUpProcess,
		Asc.c_oAscSmartArtTypes.HorizontalHierarchy,
		Asc.c_oAscSmartArtTypes.HorizontalLabeledHierarchy,
		Asc.c_oAscSmartArtTypes.HorizontalMultiLevelHierarchy,
		Asc.c_oAscSmartArtTypes.HorizontalOrganizationChart,
		Asc.c_oAscSmartArtTypes.HorizontalBulletList,
		Asc.c_oAscSmartArtTypes.ClosedChevronProcess,
		Asc.c_oAscSmartArtTypes.HierarchyList,
		Asc.c_oAscSmartArtTypes.Hierarchy,
		Asc.c_oAscSmartArtTypes.LabeledHierarchy,
		Asc.c_oAscSmartArtTypes.InvertedPyramid,
		Asc.c_oAscSmartArtTypes.CircleRelationship,
		Asc.c_oAscSmartArtTypes.CircleAccentTimeline,
		Asc.c_oAscSmartArtTypes.CircularBendingProcess,
		Asc.c_oAscSmartArtTypes.ArrowRibbon,
		Asc.c_oAscSmartArtTypes.LinearVenn,
		Asc.c_oAscSmartArtTypes.TitledMatrix,
		Asc.c_oAscSmartArtTypes.IncreasingCircleProcess,
		Asc.c_oAscSmartArtTypes.NonDirectionalCycle,
		Asc.c_oAscSmartArtTypes.ContinuousBlockProcess,
		Asc.c_oAscSmartArtTypes.ContinuousCycle,
		Asc.c_oAscSmartArtTypes.DescendingBlockList,
		Asc.c_oAscSmartArtTypes.StepDownProcess,
		Asc.c_oAscSmartArtTypes.ReverseList,
		Asc.c_oAscSmartArtTypes.OrganizationChart,
		Asc.c_oAscSmartArtTypes.NameAndTitleOrganizationChart,
		Asc.c_oAscSmartArtTypes.AlternatingFlow,
		Asc.c_oAscSmartArtTypes.PyramidList,
		Asc.c_oAscSmartArtTypes.PlusAndMinus,
		Asc.c_oAscSmartArtTypes.RepeatingBendingProcess,
		Asc.c_oAscSmartArtTypes.DetailedProcess,
		Asc.c_oAscSmartArtTypes.HalfCircleOrganizationChart,
		Asc.c_oAscSmartArtTypes.PhasedProcess,
		Asc.c_oAscSmartArtTypes.BasicVenn,
		Asc.c_oAscSmartArtTypes.BasicTimeline,
		Asc.c_oAscSmartArtTypes.BasicPie,
		Asc.c_oAscSmartArtTypes.BasicMatrix,
		Asc.c_oAscSmartArtTypes.BasicPyramid,
		Asc.c_oAscSmartArtTypes.BasicRadial,
		Asc.c_oAscSmartArtTypes.BasicTarget,
		Asc.c_oAscSmartArtTypes.BasicBlockList,
		Asc.c_oAscSmartArtTypes.BasicBendingProcess,
		Asc.c_oAscSmartArtTypes.BasicProcess,
		Asc.c_oAscSmartArtTypes.BasicChevronProcess,
		Asc.c_oAscSmartArtTypes.BasicCycle,
		Asc.c_oAscSmartArtTypes.OpposingIdeas,
		Asc.c_oAscSmartArtTypes.OpposingArrows,
		Asc.c_oAscSmartArtTypes.RandomToResultProcess,
		Asc.c_oAscSmartArtTypes.SubStepProcess,
		Asc.c_oAscSmartArtTypes.PieProcess,
		Asc.c_oAscSmartArtTypes.AccentProcess,
		Asc.c_oAscSmartArtTypes.RadialVenn,
		Asc.c_oAscSmartArtTypes.RadialCycle,
		Asc.c_oAscSmartArtTypes.RadialCluster,
		Asc.c_oAscSmartArtTypes.RadialList,
		Asc.c_oAscSmartArtTypes.MultiDirectionalCycle,
		Asc.c_oAscSmartArtTypes.DivergingRadial,
		Asc.c_oAscSmartArtTypes.DivergingArrows,
		Asc.c_oAscSmartArtTypes.GroupedList,
		Asc.c_oAscSmartArtTypes.SegmentedPyramid,
		Asc.c_oAscSmartArtTypes.SegmentedProcess,
		Asc.c_oAscSmartArtTypes.SegmentedCycle,
		Asc.c_oAscSmartArtTypes.GridMatrix,
		Asc.c_oAscSmartArtTypes.StackedList,
		Asc.c_oAscSmartArtTypes.ProcessList,
		Asc.c_oAscSmartArtTypes.SquareAccentList,
		Asc.c_oAscSmartArtTypes.LinedList,
		Asc.c_oAscSmartArtTypes.ContinuousArrowProcess,
		Asc.c_oAscSmartArtTypes.CircleArrowProcess,
		Asc.c_oAscSmartArtTypes.ProcessArrows,
		Asc.c_oAscSmartArtTypes.StaggeredProcess,
		Asc.c_oAscSmartArtTypes.ConvergingRadial,
		Asc.c_oAscSmartArtTypes.ConvergingArrows,
		Asc.c_oAscSmartArtTypes.TableHierarchy,
		Asc.c_oAscSmartArtTypes.TableList,
		Asc.c_oAscSmartArtTypes.TextCycle,
		Asc.c_oAscSmartArtTypes.TrapezoidList,
		Asc.c_oAscSmartArtTypes.DescendingProcess,
		Asc.c_oAscSmartArtTypes.ChevronList,
		Asc.c_oAscSmartArtTypes.Equation,
		Asc.c_oAscSmartArtTypes.CounterbalanceArrows,
		Asc.c_oAscSmartArtTypes.TargetList,
		Asc.c_oAscSmartArtTypes.CycleMatrix,
		Asc.c_oAscSmartArtTypes.AlternatingHexagonList,
		Asc.c_oAscSmartArtTypes.Gear,

		Asc.c_oAscSmartArtTypes.ArchitectureLayout,
		Asc.c_oAscSmartArtTypes.ChevronAccentProcess,
		Asc.c_oAscSmartArtTypes.CircleProcess,
		Asc.c_oAscSmartArtTypes.ConvergingText,
		Asc.c_oAscSmartArtTypes.HexagonRadial,
		Asc.c_oAscSmartArtTypes.InterconnectedBlockProcess,
		Asc.c_oAscSmartArtTypes.InterconnectedRings,
		Asc.c_oAscSmartArtTypes.TabList,
		Asc.c_oAscSmartArtTypes.TabbedArc,
		Asc.c_oAscSmartArtTypes.VaryingWidthList,
		Asc.c_oAscSmartArtTypes.VerticalBracketList
	];
	const arrSmartArtTypesWithImagePlaceholders = [
		Asc.c_oAscSmartArtTypes.AccentedPicture,
		Asc.c_oAscSmartArtTypes.CircularPictureCallout,
		Asc.c_oAscSmartArtTypes.PictureCaptionList,
		Asc.c_oAscSmartArtTypes.RadialPictureList,
		Asc.c_oAscSmartArtTypes.SnapshotPictureList,
		Asc.c_oAscSmartArtTypes.SpiralPicture,
		Asc.c_oAscSmartArtTypes.CaptionedPictures,
		Asc.c_oAscSmartArtTypes.BendingPictureCaption,
		Asc.c_oAscSmartArtTypes.PictureFrame,
		Asc.c_oAscSmartArtTypes.BendingPictureSemiTransparentText,
		Asc.c_oAscSmartArtTypes.BendingPictureBlocks,
		Asc.c_oAscSmartArtTypes.BendingPictureCaptionList,
		Asc.c_oAscSmartArtTypes.TitledPictureBlocks,
		Asc.c_oAscSmartArtTypes.PictureGrid,
		Asc.c_oAscSmartArtTypes.PictureAccentBlocks,
		Asc.c_oAscSmartArtTypes.PictureStrips,
		Asc.c_oAscSmartArtTypes.ThemePictureAccent,
		Asc.c_oAscSmartArtTypes.ThemePictureGrid,
		Asc.c_oAscSmartArtTypes.ThemePictureAlternatingAccent,
		Asc.c_oAscSmartArtTypes.TitledPictureAccentList,
		Asc.c_oAscSmartArtTypes.AlternatingPictureBlocks,
		Asc.c_oAscSmartArtTypes.AscendingPictureAccentProcess,
		Asc.c_oAscSmartArtTypes.AlternatingPictureCircles,
		Asc.c_oAscSmartArtTypes.TitlePictureLineup,
		Asc.c_oAscSmartArtTypes.PictureLineup,
		Asc.c_oAscSmartArtTypes.FramedTextPicture,
		Asc.c_oAscSmartArtTypes.HexagonCluster,
		Asc.c_oAscSmartArtTypes.BubblePictureList,
		Asc.c_oAscSmartArtTypes.CirclePictureHierarchy,
		Asc.c_oAscSmartArtTypes.HorizontalPictureList,
		Asc.c_oAscSmartArtTypes.ContinuousPictureList,
		Asc.c_oAscSmartArtTypes.VerticalPictureList,
		Asc.c_oAscSmartArtTypes.VerticalPictureAccentList,
		Asc.c_oAscSmartArtTypes.BendingPictureAccentList,
		Asc.c_oAscSmartArtTypes.PictureAccentList,
		Asc.c_oAscSmartArtTypes.PictureAccentProcess,
		Asc.c_oAscSmartArtTypes.PictureOrganizationChart,
	];

	function getSmartArtByType(nSmartArtType)
	{
		const oSmartArt = new AscFormat.SmartArt();
		oSmartArt.fillByPreset(nSmartArtType);
		return oSmartArt;
	}

	AscTest.CreateLogicDocument();
	AscCommon.g_oBinarySmartArts.checkLoadDrawing().then(function ()
	{
		const arrPromises = [];
		for (let sSmartArtType in Asc.c_oAscSmartArtTypes)
		{
			arrPromises.push(AscCommon.g_oBinarySmartArts.checkLoadData(Asc.c_oAscSmartArtTypes[sSmartArtType]));
		}
		return Promise.all(arrPromises);
	}).then(function ()
	{
		startTests();
	});

	QUnit.module("Test truth smart art image placeholders");

	function startTests()
	{
		QUnit.start();
		QUnit.test("Test smartarts without image placeholders", function (assert)
		{
			for (let i = 0; i < arrSmartArtTypesWithoutImagePlaceholders.length; i += 1)
			{
				const nSmartArtType = arrSmartArtTypesWithoutImagePlaceholders[i];
				const oSmartArt = new AscFormat.SmartArt();
				oSmartArt.fillByPreset(nSmartArtType);
				let bHaveImagePlaceholder = false;
				oSmartArt.spTree[0].spTree.forEach(function (oShape)
				{
					bHaveImagePlaceholder = bHaveImagePlaceholder || !!oShape.isActiveBlipFillPlaceholder();
				});
				assert.strictEqual(bHaveImagePlaceholder, false, 'Test smartart with ' + nSmartArtType + ' nType');
			}
		});

		const oAnswersSmartArtWithPlaceholders = {};
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.AccentedPicture] = [0, 2, 4, 6];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.CircularPictureCallout] = [3, 5, 7, 9];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.PictureCaptionList] = [0, 2, 4, 6];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.RadialPictureList] = [2, 4, 6];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.SnapshotPictureList] = [2];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.SpiralPicture] = [0, 2, 5, 9, 14];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.CaptionedPictures] = [1, 5, 9];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.BendingPictureCaption] = [0, 2];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.PictureFrame] = [0, 1, 2];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.BendingPictureSemiTransparentText] = [0, 2, 4];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.BendingPictureBlocks] = [0, 2, 4];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.BendingPictureCaptionList] = [0, 2, 4];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.TitledPictureBlocks] = [0, 3, 6];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.PictureGrid] = [1, 3, 5, 7];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.PictureAccentBlocks] = [0, 2, 4];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.PictureStrips] = [1, 3, 5];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.ThemePictureAccent] = [0, 3, 6, 9, 12, 15];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.ThemePictureGrid] = [0, 3, 5, 7, 9];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.ThemePictureAlternatingAccent] = [0, 2, 4];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.TitledPictureAccentList] = [1, 3];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.AlternatingPictureBlocks] = [1, 3, 5];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.AscendingPictureAccentProcess] = [11, 13];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.AlternatingPictureCircles] = [1, 5, 9];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.TitlePictureLineup] = [1, 5, 9];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.PictureLineup] = [0, 3, 6];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.FramedTextPicture] = [0];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.HexagonCluster] = [2, 6, 10];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.BubblePictureList] = [2, 4, 8];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.CirclePictureHierarchy] = [5, 7, 9, 11, 13, 15];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.HorizontalPictureList] = [1, 3, 5];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.ContinuousPictureList] = [1, 3, 5];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.VerticalPictureList] = [1, 3, 5];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.VerticalPictureAccentList] = [1, 3, 5];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.BendingPictureAccentList] = [2, 5, 8];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.PictureAccentList] = [2, 5, 8];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.PictureAccentProcess] = [0, 3, 6];
		oAnswersSmartArtWithPlaceholders[Asc.c_oAscSmartArtTypes.PictureOrganizationChart] = [5, 7, 9, 11, 13];

		QUnit.test("Test smartarts with image placeholders", function (assert)
		{
			function compareResultArray(nSmartArtType)
			{
				const arrAnswers = oAnswersSmartArtWithPlaceholders[nSmartArtType];
				const arrResult = [];
				const oSmartArt = getSmartArtByType(nSmartArtType);
				const arrSpTree = oSmartArt.spTree[0].spTree;
				for (let i = 0; i < arrSpTree.length; i += 1)
				{
					const oShape = arrSpTree[i];
					if (oShape.isActiveBlipFillPlaceholder())
					{
						arrResult.push(i);
					}
				}
				assert.deepEqual(arrResult, arrAnswers, 'Check answers for smartart with image placeholders. ' + nSmartArtType + ' nType');
			}

			for (let i = 0; i < arrSmartArtTypesWithImagePlaceholders.length; i++)
			{
				const nSmartArtType = arrSmartArtTypesWithImagePlaceholders[i];
				compareResultArray(nSmartArtType);
			}
		});
	}

});
