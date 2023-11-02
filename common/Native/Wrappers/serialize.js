/*
 * (c) Copyright Ascensio System SIA 2010-2023
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

// FONTS
function asc_menu_ReadFontFamily(_params, _cursor)
{
	var _fontfamily = { Name : undefined, Index : -1 };
	var _continue = true;
	while (_continue)
	{
		var _attr = _params[_cursor.pos++];
		switch (_attr)
		{
			case 0:
			{
				_fontfamily.Name = _params[_cursor.pos++];
				break;
			}
			case 1:
			{
				_fontfamily.Index = _params[_cursor.pos++];
				break;
			}
			case 255:
			default:
			{
				_continue = false;
				break;
			}
		}
	}
	return _fontfamily;
}
function asc_menu_WriteFontFamily(_type, _family, _stream)
{
	if (!_family)
		return;

	_stream["WriteByte"](_type);

	if (_family.Name !== undefined && _family.Name !== null)
	{
		_stream["WriteByte"](0);
		_stream["WriteString2"](_family.Name);
	}
	if (_family.Index !== undefined && _family.Index !== null)
	{
		_stream["WriteByte"](1);
		_stream["WriteLong"](_family.Index);
	}

	_stream["WriteByte"](255);
}

// PARAFRAME
function asc_menu_ReadParaFrame(_params, _cursor)
{
	var _frame = new Asc.asc_CParagraphFrame();
	var _continue = true;
	while (_continue)
	{
		var _attr = _params[_cursor.pos++];
		switch (_attr)
		{
			case 0:
			{
				_frame.FromDropCapMenu = _params[_cursor.pos++];
				break;
			}
			case 1:
			{
				_frame.DropCap = _params[_cursor.pos++];
				break;
			}
			case 2:
			{
				_frame.W = _params[_cursor.pos++];
				break;
			}
			case 3:
			{
				_frame.H = _params[_cursor.pos++];
				break;
			}
			case 4:
			{
				_frame.HAnchor = _params[_cursor.pos++];
				break;
			}
			case 5:
			{
				_frame.HRule = _params[_cursor.pos++];
				break;
			}
			case 6:
			{
				_frame.HSpace = _params[_cursor.pos++];
				break;
			}
			case 7:
			{
				_frame.VAnchor = _params[_cursor.pos++];
				break;
			}
			case 8:
			{
				_frame.VSpace = _params[_cursor.pos++];
				break;
			}
			case 9:
			{
				_frame.X = _params[_cursor.pos++];
				break;
			}
			case 10:
			{
				_frame.Y = _params[_cursor.pos++];
				break;
			}
			case 11:
			{
				_frame.XAlign = _params[_cursor.pos++];
				break;
			}
			case 12:
			{
				_frame.YAlign = _params[_cursor.pos++];
				break;
			}
			case 13:
			{
				_frame.Lines = _params[_cursor.pos++];
				break;
			}
			case 14:
			{
				_frame.Wrap = _params[_cursor.pos++];
				break;
			}
			case 15:
			{
				_frame.Brd = asc_menu_ReadParaBorders(_params, _cursor);
				break;
			}
			case 16:
			{
				_frame.Shd = asc_menu_ReadParaShd(_params, _cursor);
				break;
			}
			case 17:
			{
				_frame.FontFamily = asc_menu_ReadFontFamily(_params, _cursor);
				break;
			}
			case 255:
			default:
			{
				_continue = false;
				break;
			}
		}
	}
	return _frame;
}
function asc_menu_WriteParaFrame(_type, _frame, _stream)
{
	if (!_frame)
		return;

	_stream["WriteByte"](_type);

	if (_frame.FromDropCapMenu !== undefined && _frame.FromDropCapMenu !== null)
	{
		_stream["WriteByte"](0);
		_stream["WriteBool"](_frame.FromDropCapMenu);
	}
	if (_frame.DropCap !== undefined && _frame.DropCap !== null)
	{
		_stream["WriteByte"](1);
		_stream["WriteLong"](_frame.DropCap);
	}
	if (_frame.W !== undefined && _frame.W !== null)
	{
		_stream["WriteByte"](2);
		_stream["WriteDouble2"](_frame.W);
	}
	if (_frame.H !== undefined && _frame.H !== null)
	{
		_stream["WriteByte"](3);
		_stream["WriteDouble2"](_frame.H);
	}
	if (_frame.HAlign !== undefined && _frame.HAlign !== null)
	{
		_stream["WriteByte"](4);
		_stream["WriteLong"](_frame.HAlign);
	}
	if (_frame.HRule !== undefined && _frame.HRule !== null)
	{
		_stream["WriteByte"](5);
		_stream["WriteLong"](_frame.HRule);
	}
	if (_frame.HSpace !== undefined && _frame.HSpace !== null)
	{
		_stream["WriteByte"](6);
		_stream["WriteDouble2"](_frame.HSpace);
	}
	if (_frame.VAnchor !== undefined && _frame.VAnchor !== null)
	{
		_stream["WriteByte"](7);
		_stream["WriteLong"](_frame.VAnchor);
	}
	if (_frame.VSpace !== undefined && _frame.VSpace !== null)
	{
		_stream["WriteByte"](8);
		_stream["WriteDouble2"](_frame.VSpace);
	}
	if (_frame.X !== undefined && _frame.X !== null)
	{
		_stream["WriteByte"](9);
		_stream["WriteDouble2"](_frame.X);
	}
	if (_frame.Y !== undefined && _frame.Y !== null)
	{
		_stream["WriteByte"](10);
		_stream["WriteDouble2"](_frame.Y);
	}
	if (_frame.XAlign !== undefined && _frame.XAlign !== null)
	{
		_stream["WriteByte"](11);
		_stream["WriteLong"](_frame.XAlign);
	}
	if (_frame.YAlign !== undefined && _frame.YAlign !== null)
	{
		_stream["WriteByte"](12);
		_stream["WriteLong"](_frame.YAlign);
	}
	if (_frame.Lines !== undefined && _frame.Lines !== null)
	{
		_stream["WriteByte"](13);
		_stream["WriteLong"](_frame.Lines);
	}
	if (_frame.Wrap !== undefined && _frame.Wrap !== null)
	{
		_stream["WriteByte"](14);
		_stream["WriteLong"](_frame.Wrap);
	}

	asc_menu_WriteParaBorders(15, _frame.Brd, _stream);
	asc_menu_WriteParaShd(16, _frame.Shd, _stream);
	asc_menu_WriteFontFamily(17, _frame.FontFamily, _stream);

	_stream["WriteByte"](255);
}

// COLOR
function asc_menu_WriteColor(_type, _color, _stream)
{
	if (!_color)
		return;

	// TODO:
	if (_color.write)
		_color.write(_type, _stream);
	else
		Asc.asc_CColor.prototype.write.call(_color, _type, _stream);
}

// MATH
function asc_menu_WriteMath(oMath, s)
{
	s["WriteLong"](oMath.Type);
	s["WriteLong"](oMath.Action);
	s["WriteBool"](oMath.CanIncreaseArgumentSize);
	s["WriteBool"](oMath.CanDecreaseArgumentSize);
	s["WriteBool"](oMath.CanInsertForcedBreak);
	s["WriteBool"](oMath.CanDeleteForcedBreak);
	s["WriteBool"](oMath.CanAlignToCharacter);
}

// POSITIONS
function asc_menu_ReadPosition(_params, _cursor)
{
	var _position = new Asc.CPosition();
	var _continue = true;
	while (_continue)
	{
		var _attr = _params[_cursor.pos++];
		switch (_attr)
		{
			case 0:
			{
				_position.X = _params[_cursor.pos++];
				break;
			}
			case 1:
			{
				_position.Y = _params[_cursor.pos++];
				break;
			}
			case 255:
			default:
			{
				_continue = false;
				break;
			}
		}
	}
	return _position;
}
function asc_menu_WritePosition(_type, _position, _stream)
{
	if (!_position)
		return;

	_stream["WriteByte"](_type);

	if (_position.X !== undefined && _position.X !== null)
	{
		_stream["WriteByte"](0);
		_stream["WriteDouble2"](_position.X);
	}
	if (_position.Y !== undefined && _position.Y !== null)
	{
		_stream["WriteByte"](1);
		_stream["WriteDouble2"](_position.Y);
	}

	_stream["WriteByte"](255);
}

function asc_menu_ReadImagePosition(_params, _cursor)
{
	var _position = new Asc.CPosition();
	var _continue = true;
	while (_continue)
	{
		var _attr = _params[_cursor.pos++];
		switch (_attr)
		{
			case 0:
			{
				_position.RelativeFrom = _params[_cursor.pos++];
				break;
			}
			case 1:
			{
				_position.UseAlign = _params[_cursor.pos++];
				break;
			}
			case 2:
			{
				_position.Align = _params[_cursor.pos++];
				break;
			}
			case 3:
			{
				_position.Value = _params[_cursor.pos++];
				break;
			}
			case 255:
			default:
			{
				_continue = false;
				break;
			}
		}
	}
	return _position;
}
function asc_menu_WriteImagePosition(_type, _position, _stream)
{
	if (!_position)
		return;

	_stream["WriteByte"](_type);

	if (_position.RelativeFrom !== undefined && _position.RelativeFrom !== null)
	{
		_stream["WriteByte"](0);
		_stream["WriteLong"](_position.RelativeFrom);
	}
	if (_position.UseAlign !== undefined && _position.UseAlign !== null)
	{
		_stream["WriteByte"](1);
		_stream["WriteBool"](_position.UseAlign);
	}
	if (_position.Align !== undefined && _position.Align !== null)
	{
		_stream["WriteByte"](2);
		_stream["WriteLong"](_position.Align);
	}
	if (_position.Value !== undefined && _position.Value !== null)
	{
		_stream["WriteByte"](3);
		_stream["WriteLong"](_position.Value);
	}

	_stream["WriteByte"](255);
}

// CHART PR
function asc_menu_ReadChartPr(_params, _cursor)
{
	var _settings = new Asc.asc_ChartSettings();
	_settings.read(_params, _cursor);
	return _settings;
}
function asc_menu_WriteChartPr(_type, _chartPr, _stream)
{
	if (!_chartPr)
		return;
	_chartPr.write(_type, _stream);
}

// SHAPE PR
function asc_menu_ReadShapePr(_params, _cursor)
{
	const _settings = new Asc.asc_CShapeProperty();
	_settings.read(_params, _cursor);
	return _settings;
}
function asc_menu_WriteShapePr(_type, _shapePr, _stream)
{
	if (!_shapePr)
		return;
	_shapePr.write(_type, _stream);
}

// IMAGE PR
function asc_menu_WriteImagePr(_imagePr, _stream)
{
	if (_imagePr.CanBeFlow !== undefined && _imagePr.CanBeFlow !== null)
	{
		_stream["WriteByte"](0);
		_stream["WriteBool"](_imagePr.CanBeFlow);
	}
	if (_imagePr.Width !== undefined && _imagePr.Width !== null)
	{
		_stream["WriteByte"](1);
		_stream["WriteDouble2"](_imagePr.Width);
	}
	if (_imagePr.Height !== undefined && _imagePr.Height !== null)
	{
		_stream["WriteByte"](2);
		_stream["WriteDouble2"](_imagePr.Height);
	}
	if (_imagePr.WrappingStyle !== undefined && _imagePr.WrappingStyle !== null)
	{
		_stream["WriteByte"](3);
		_stream["WriteLong"](_imagePr.WrappingStyle);
	}

	if(_imagePr.Paddings) {
		_imagePr.Paddings.write(4, _stream);
	}
	asc_menu_WritePosition(5, _imagePr.Position, _stream);

	if (_imagePr.AllowOverlap !== undefined && _imagePr.AllowOverlap !== null)
	{
		_stream["WriteByte"](6);
		_stream["WriteBool"](_imagePr.AllowOverlap);
	}

	asc_menu_WriteImagePosition(7, _imagePr.PositionH, _stream);
	asc_menu_WriteImagePosition(8, _imagePr.PositionV, _stream);

	if (_imagePr.Internal_Position !== undefined && _imagePr.Internal_Position !== null)
	{
		_stream["WriteByte"](9);
		_stream["WriteLong"](_imagePr.Internal_Position);
	}

	if (_imagePr.ImageUrl !== undefined && _imagePr.ImageUrl !== null)
	{
		_stream["WriteByte"](10);
		_stream["WriteString2"](_imagePr.ImageUrl);
	}

	if (_imagePr.Locked !== undefined && _imagePr.Locked !== null)
	{
		_stream["WriteByte"](11);
		_stream["WriteBool"](_imagePr.Locked);
	}

	asc_menu_WriteChartPr(12, _imagePr.ChartProperties, _stream);
	asc_menu_WriteShapePr(13, _imagePr.ShapeProperties, _stream);

	if (_imagePr.ChangeLevel !== undefined && _imagePr.ChangeLevel !== null)
	{
		_stream["WriteByte"](14);
		_stream["WriteLong"](_imagePr.ChangeLevel);
	}

	if (_imagePr.Group !== undefined && _imagePr.Group !== null)
	{
		_stream["WriteByte"](15);
		_stream["WriteLong"](_imagePr.Group);
	}

	if (_imagePr.fromGroup !== undefined && _imagePr.fromGroup !== null)
	{
		_stream["WriteByte"](16);
		_stream["WriteBool"](_imagePr.fromGroup);
	}
	if (_imagePr.severalCharts !== undefined && _imagePr.severalCharts !== null)
	{
		_stream["WriteByte"](17);
		_stream["WriteBool"](_imagePr.severalCharts);
	}

	if (_imagePr.severalChartTypes !== undefined && _imagePr.severalChartTypes !== null)
	{
		_stream["WriteByte"](18);
		_stream["WriteLong"](_imagePr.severalChartTypes);
	}
	if (_imagePr.severalChartStyles !== undefined && _imagePr.severalChartStyles !== null)
	{
		_stream["WriteByte"](19);
		_stream["WriteLong"](_imagePr.severalChartStyles);
	}
	if (_imagePr.verticalTextAlign !== undefined && _imagePr.verticalTextAlign !== null)
	{
		_stream["WriteByte"](20);
		_stream["WriteLong"](_imagePr.verticalTextAlign);
	}
	if (_imagePr.vert !== undefined && _imagePr.vert !== null) {
		_stream["WriteByte"](21);
		_stream["WriteLong"](_imagePr.vert);
	}
	_stream["WriteByte"](255);
}

// PARAGRAPH PR
function asc_menu_ReadParaInd(_params, _cursor)
{
	var _ind = new CParaInd();
	var _continue = true;
	while (_continue)
	{
		var _attr = _params[_cursor.pos++];
		switch (_attr)
		{
			case 0:
			{
				_ind.Left = _params[_cursor.pos++];
				break;
			}
			case 1:
			{
				_ind.Right = _params[_cursor.pos++];
				break;
			}
			case 2:
			{
				_ind.FirstLine = _params[_cursor.pos++];
				break;
			}
			case 255:
			default:
			{
				_continue = false;
				break;
			}
		}
	}
	return _ind;
}
function asc_menu_WriteParaInd(_type, _ind, _stream)
{
	if (!_ind)
		return;

	_stream["WriteByte"](_type);

	if (_ind.Left !== undefined && _ind.Left !== null)
	{
		_stream["WriteByte"](0);
		_stream["WriteDouble2"](_ind.Left);
	}
	if (_ind.Right !== undefined && _ind.Right !== null)
	{
		_stream["WriteByte"](1);
		_stream["WriteDouble2"](_ind.Right);
	}
	if (_ind.FirstLine !== undefined && _ind.FirstLine !== null)
	{
		_stream["WriteByte"](2);
		_stream["WriteDouble2"](_ind.FirstLine);
	}

	_stream["WriteByte"](255);
}

function asc_menu_ReadParaSpacing(_params, _cursor)
{
	var _spacing = new CParaSpacing();
	var _continue = true;
	while (_continue)
	{
		var _attr = _params[_cursor.pos++];
		switch (_attr)
		{
			case 0:
			{
				_spacing.Line = _params[_cursor.pos++];
				break;
			}
			case 1:
			{
				_spacing.LineRule = _params[_cursor.pos++];
				break;
			}
			case 2:
			{
				_spacing.Before = _params[_cursor.pos++];
				break;
			}
			case 3:
			{
				_spacing.BeforeAutoSpacing = _params[_cursor.pos++];
				break;
			}
			case 4:
			{
				_spacing.After = _params[_cursor.pos++];
				break;
			}
			case 5:
			{
				_spacing.AfterAutoSpacing = _params[_cursor.pos++];
				break;
			}
			case 255:
			default:
			{
				_continue = false;
				break;
			}
		}
	}
	return _spacing;
}
function asc_menu_WriteParaSpacing(_type, _spacing, _stream)
{
	if (!_spacing)
		return;

	_stream["WriteByte"](_type);

	if (_spacing.Line !== undefined && _spacing.Line !== null)
	{
		_stream["WriteByte"](0);
		_stream["WriteDouble2"](_spacing.Line);
	}
	if (_spacing.LineRule !== undefined && _spacing.LineRule !== null)
	{
		_stream["WriteByte"](1);
		_stream["WriteLong"](_spacing.LineRule);
	}
	if (_spacing.Before !== undefined && _spacing.Before !== null)
	{
		_stream["WriteByte"](2);
		_stream["WriteDouble2"](_spacing.Before);
	}
	if (_spacing.BeforeAutoSpacing !== undefined && _spacing.BeforeAutoSpacing !== null)
	{
		_stream["WriteByte"](3);
		_stream["WriteBool"](_spacing.BeforeAutoSpacing);
	}
	if (_spacing.After !== undefined && _spacing.After !== null)
	{
		_stream["WriteByte"](4);
		_stream["WriteDouble2"](_spacing.After);
	}
	if (_spacing.AfterAutoSpacing !== undefined && _spacing.AfterAutoSpacing !== null)
	{
		_stream["WriteByte"](5);
		_stream["WriteBool"](_spacing.AfterAutoSpacing);
	}

	_stream["WriteByte"](255);
}

function asc_menu_ReadParaBorder(_params, _cursor)
{
	var _border = new Asc.asc_CTextBorder();
	var _continue = true;
	while (_continue)
	{
		var _attr = _params[_cursor.pos++];
		switch (_attr)
		{
			case 0:
			{
				_border.Color = AscCommon.asc_menu_ReadColor(_params, _cursor);
				break;
			}
			case 1:
			{
				_border.Size = _params[_cursor.pos++];
				break;
			}
			case 2:
			{
				_border.Value = _params[_cursor.pos++];
				break;
			}
			case 3:
			{
				_border.Space = _params[_cursor.pos++];
				break;
			}
			case 255:
			default:
			{
				_continue = false;
				break;
			}
		}
	}
	return _border;
}
function asc_menu_ReadParaBorders(_params, _cursor)
{
	var _border = new Asc.asc_CParagraphBorders();
	var _continue = true;
	while (_continue)
	{
		var _attr = _params[_cursor.pos++];
		switch (_attr)
		{
			case 0:
			{
				_border.Left = asc_menu_ReadParaBorder(_params, _cursor);
				break;
			}
			case 1:
			{
				_border.Top = asc_menu_ReadParaBorder(_params, _cursor);
				break;
			}
			case 2:
			{
				_border.Right = asc_menu_ReadParaBorder(_params, _cursor);
				break;
			}
			case 3:
			{
				_border.Bottom = asc_menu_ReadParaBorder(_params, _cursor);
				break;
			}
			case 4:
			{
				_border.Between = asc_menu_ReadParaBorder(_params, _cursor);
				break;
			}
			case 255:
			default:
			{
				_continue = false;
				break;
			}
		}
	}
	return _border;
}
function asc_menu_WriteParaBorder(_type, _border, _stream)
{
	if (!_border)
		return;

	_stream["WriteByte"](_type);

	asc_menu_WriteColor(0, _border.Color, _stream);

	if (_border.Size !== undefined && _border.Size !== null)
	{
		_stream["WriteByte"](1);
		_stream["WriteDouble2"](_border.Size);
	}
	if (_border.Value !== undefined && _border.Value !== null)
	{
		_stream["WriteByte"](2);
		_stream["WriteLong"](_border.Value);
	}
	if (_border.Space !== undefined && _border.Space !== null)
	{
		_stream["WriteByte"](3);
		_stream["WriteDouble2"](_border.Space);
	}

	_stream["WriteByte"](255);
}
function asc_menu_WriteParaBorders(_type, _borders, _stream)
{
	if (!_borders)
		return;

	_stream["WriteByte"](_type);

	asc_menu_WriteParaBorder(0, _borders.Left, _stream);
	asc_menu_WriteParaBorder(1, _borders.Top, _stream);
	asc_menu_WriteParaBorder(2, _borders.Right, _stream);
	asc_menu_WriteParaBorder(3, _borders.Bottom, _stream);
	asc_menu_WriteParaBorder(4, _borders.Between, _stream);

	_stream["WriteByte"](255);
}

function asc_menu_ReadParaShd(_params, _cursor)
{
	var _shd = new Asc.asc_CParagraphShd();
	var _continue = true;
	while (_continue)
	{
		var _attr = _params[_cursor.pos++];
		switch (_attr)
		{
			case 0:
			{
				_shd.Value = _params[_cursor.pos++];
				break;
			}
			case 1:
			{
				_shd.Color = AscCommon.asc_menu_ReadColor(_params, _cursor);
				break;
			}
			case 255:
			default:
			{
				_continue = false;
				break;
			}
		}
	}
	if(_shd.Value === Asc.c_oAscShd.Clear)
	{
		if(_shd.Color)
		{
			if(_shd.Color.Auto)
			{
				_shd.Color.r = 0;
				_shd.Color.g = 0;
				_shd.Color.b = 0;
				_shd.Fill = {};
				_shd.Fill.Auto = true;
				_shd.Fill.r = 255;
				_shd.Fill.g = 255;
				_shd.Fill.b = 255;
			}
			else
			{
				_shd.Color.Auto = false;
				_shd.Fill.Auto = false;
				_shd.Fill.r = _shd.Color.r;
				_shd.Fill.g = _shd.Color.g;
				_shd.Fill.b = _shd.Color.b;
				var Unifill        = new AscFormat.CUniFill();
				Unifill.fill       = new AscFormat.CSolidFill();
				Unifill.fill.color = AscFormat.CorrectUniColor(_shd.Color, Unifill.fill.color, 1);
				_shd.Unifill = Unifill;
			}
		}
	}
	return _shd;
}
function asc_menu_WriteParaShd(_type, _shd, _stream)
{
	if (!_shd)
		return;

	_stream["WriteByte"](_type);

	if (_shd.Value !== undefined && _shd.Value !== null)
	{
		_stream["WriteByte"](0);
		_stream["WriteLong"](_shd.Value);
	}

	asc_menu_WriteColor(1, _shd.Color, _stream);

	_stream["WriteByte"](255);
}

function asc_menu_ReadParaTabs(_params, _cursor)
{
	var _tabs = new Asc.asc_CParagraphTabs();

	var _count = _params[_cursor.pos++];

	for (var i = 0; i < _count; i++)
	{
		var _tab = new Asc.asc_CParagraphTab();
		var _continue = true;
		while (_continue)
		{
			var _attr = _params[_cursor.pos++];
			switch (_attr)
			{
				case 0:
				{
					_tab.Pos = _params[_cursor.pos++];
					break;
				}
				case 1:
				{
					_tab.Value = _params[_cursor.pos++];
					break;
				}
				case 255:
				default:
				{
					_continue = false;
					break;
				}
			}
		}

		_tabs.Tabs.push(_tab);
	}
	return _tabs;
}
function asc_menu_WriteParaTabs(_type, _tabs, _stream)
{
	if (!_tabs)
		return;

	_stream["WriteByte"](_type);

	var _len = _tabs.Tabs.length;
	_stream["WriteLong"](_len);

	for (var i = 0; i < _len; i++)
	{
		if (_tabs.Tabs[i].Pos !== undefined && _tabs.Tabs[i].Pos !== null)
		{
			_stream["WriteByte"](0);
			_stream["WriteDouble2"](_tabs.Tabs[i].Pos);
		}
		if (_tabs.Tabs[i].Value !== undefined && _tabs.Tabs[i].Value !== null)
		{
			_stream["WriteByte"](1);
			_stream["WriteLong"](_tabs.Tabs[i].Value);
		}
		_stream["WriteByte"](255);
	}
}

function asc_menu_ReadParaListType(_params, _cursor)
{
	var _list = new AscCommon.asc_CListType();
	var _continue = true;
	while (_continue)
	{
		var _attr = _params[_cursor.pos++];
		switch (_attr)
		{
			case 0:
			{
				_list.Type = _params[_cursor.pos++];
				break;
			}
			case 1:
			{
				_list.SubType = _params[_cursor.pos++];
				break;
			}
			case 255:
			default:
			{
				_continue = false;
				break;
			}
		}
	}
	return _list;
}
function asc_menu_WriteParaListType(_type, _list, _stream)
{
	if (!_list)
		return;

	_stream["WriteByte"](_type);

	if (_list.Type !== undefined && _list.Type !== null)
	{
		_stream["WriteByte"](0);
		_stream["WriteLong"](_list.Type);
	}
	if (_list.SubType !== undefined && _list.SubType !== null)
	{
		_stream["WriteByte"](1);
		_stream["WriteLong"](_list.SubType);
	}

	_stream["WriteByte"](255);
}

function asc_menu_WriteParagraphPr(_paraPr, _stream)
{
	if (_paraPr.ContextualSpacing !== undefined && _paraPr.ContextualSpacing !== null)
	{
		_stream["WriteByte"](0);
		_stream["WriteBool"](_paraPr.ContextualSpacing);
	}
	asc_menu_WriteParaInd(1, _paraPr.Ind, _stream);

	if (_paraPr.KeepLines !== undefined && _paraPr.KeepLines !== null)
	{
		_stream["WriteByte"](2);
		_stream["WriteBool"](_paraPr.KeepLines);
	}
	if (_paraPr.KeepNext !== undefined && _paraPr.KeepNext !== null)
	{
		_stream["WriteByte"](3);
		_stream["WriteBool"](_paraPr.KeepNext);
	}
	if (_paraPr.WidowControl !== undefined && _paraPr.WidowControl !== null)
	{
		_stream["WriteByte"](4);
		_stream["WriteBool"](_paraPr.WidowControl);
	}
	if (_paraPr.PageBreakBefore !== undefined && _paraPr.PageBreakBefore !== null)
	{
		_stream["WriteByte"](5);
		_stream["WriteBool"](_paraPr.PageBreakBefore);
	}

	asc_menu_WriteParaSpacing(6, _paraPr.Spacing, _stream);
	asc_menu_WriteParaBorders(7, _paraPr.Brd, _stream);
	asc_menu_WriteParaShd(8, _paraPr.Shd, _stream);

	if (_paraPr.Locked !== undefined && _paraPr.Locked !== null)
	{
		_stream["WriteByte"](9);
		_stream["WriteBool"](_paraPr.Locked);
	}
	if (_paraPr.CanAddTable !== undefined && _paraPr.CanAddTable !== null)
	{
		_stream["WriteByte"](10);
		_stream["WriteBool"](_paraPr.CanAddTable);
	}
	if (_paraPr.CanAddDropCap !== undefined && _paraPr.CanAddDropCap !== null)
	{
		_stream["WriteByte"](11);
		_stream["WriteBool"](_paraPr.CanAddDropCap);
	}

	if (_paraPr.DefaultTab !== undefined && _paraPr.DefaultTab !== null)
	{
		_stream["WriteByte"](12);
		_stream["WriteDouble2"](_paraPr.DefaultTab);
	}

	asc_menu_WriteParaTabs(13, _paraPr.Tabs, _stream);
	asc_menu_WriteParaFrame(14, _paraPr.FramePr, _stream);

	if (_paraPr.Subscript !== undefined && _paraPr.Subscript !== null)
	{
		_stream["WriteByte"](15);
		_stream["WriteBool"](_paraPr.Subscript);
	}
	if (_paraPr.Superscript !== undefined && _paraPr.Superscript !== null)
	{
		_stream["WriteByte"](16);
		_stream["WriteBool"](_paraPr.Superscript);
	}
	if (_paraPr.SmallCaps !== undefined && _paraPr.SmallCaps !== null)
	{
		_stream["WriteByte"](17);
		_stream["WriteBool"](_paraPr.SmallCaps);
	}
	if (_paraPr.AllCaps !== undefined && _paraPr.AllCaps !== null)
	{
		_stream["WriteByte"](18);
		_stream["WriteBool"](_paraPr.AllCaps);
	}
	if (_paraPr.Strikeout !== undefined && _paraPr.Strikeout !== null)
	{
		_stream["WriteByte"](19);
		_stream["WriteBool"](_paraPr.Strikeout);
	}
	if (_paraPr.DStrikeout !== undefined && _paraPr.DStrikeout !== null)
	{
		_stream["WriteByte"](20);
		_stream["WriteBool"](_paraPr.DStrikeout);
	}

	if (_paraPr.TextSpacing !== undefined && _paraPr.TextSpacing !== null)
	{
		_stream["WriteByte"](21);
		_stream["WriteDouble2"](_paraPr.TextSpacing);
	}
	if (_paraPr.Position !== undefined && _paraPr.Position !== null)
	{
		_stream["WriteByte"](22);
		_stream["WriteDouble2"](_paraPr.Position);
	}

	asc_menu_WriteParaListType(23, _paraPr.ListType, _stream);

	if (_paraPr.StyleName !== undefined && _paraPr.StyleName !== null)
	{
		_stream["WriteByte"](24);
		_stream["WriteString2"](_paraPr.StyleName);
	}

	if (_paraPr.Jc !== undefined && _paraPr.Jc !== null)
	{
		_stream["WriteByte"](25);
		_stream["WriteLong"](_paraPr.Jc);
	}

	_stream["WriteByte"](255);
}

// CELLS
function asc_menu_ReadCellMargins(_params, _cursor)
{
	var _paddings = new Asc.CMargins();
	var _continue = true;
	while (_continue)
	{
		var _attr = _params[_cursor.pos++];
		switch (_attr)
		{
			case 0:
			{
				_paddings.Left = _params[_cursor.pos++];
				break;
			}
			case 1:
			{
				_paddings.Top = _params[_cursor.pos++];
				break;
			}
			case 2:
			{
				_paddings.Right = _params[_cursor.pos++];
				break;
			}
			case 3:
			{
				_paddings.Bottom = _params[_cursor.pos++];
				break;
			}
			case 4:
			{
				_paddings.Flag = _params[_cursor.pos++];
				break;
			}
			case 255:
			default:
			{
				_continue = false;
				break;
			}
		}
	}
	return _paddings;
}
function asc_menu_WriteCellMargins(_type, _margins, _stream)
{
	if (!_margins)
		return;

	_stream["WriteByte"](_type);

	if (_margins.Left !== undefined && _margins.Left !== null)
	{
		_stream["WriteByte"](0);
		_stream["WriteDouble2"](_margins.Left);
	}
	if (_margins.Top !== undefined && _margins.Top !== null)
	{
		_stream["WriteByte"](1);
		_stream["WriteDouble2"](_margins.Top);
	}
	if (_margins.Right !== undefined && _margins.Right !== null)
	{
		_stream["WriteByte"](2);
		_stream["WriteDouble2"](_margins.Right);
	}
	if (_margins.Bottom !== undefined && _margins.Bottom !== null)
	{
		_stream["WriteByte"](3);
		_stream["WriteDouble2"](_margins.Bottom);
	}
	if (_margins.Flag !== undefined && _margins.Flag !== null)
	{
		_stream["WriteByte"](4);
		_stream["WriteLong"](_margins.Flag);
	}

	_stream["WriteByte"](255);
}

function asc_menu_ReadCellBorders(_params, _cursor)
{
	var _borders = new Asc.CBorders();
	var _continue = true;
	while (_continue)
	{
		var _attr = _params[_cursor.pos++];
		switch (_attr)
		{
			case 0:
			{
				_borders.Left = asc_menu_ReadParaBorder(_params, _cursor);
				break;
			}
			case 1:
			{
				_borders.Top = asc_menu_ReadParaBorder(_params, _cursor);
				break;
			}
			case 2:
			{
				_borders.Right = asc_menu_ReadParaBorder(_params, _cursor);
				break;
			}
			case 3:
			{
				_borders.Bottom = asc_menu_ReadParaBorder(_params, _cursor);
				break;
			}
			case 4:
			{
				_borders.InsideH = asc_menu_ReadParaBorder(_params, _cursor);
				break;
			}
			case 5:
			{
				_borders.InsideV = asc_menu_ReadParaBorder(_params, _cursor);
				break;
			}
			case 255:
			default:
			{
				_continue = false;
				break;
			}
		}
	}
	return _borders;
}
function asc_menu_WriteCellBorders(_type, _borders, _stream)
{
	if (!_borders)
		return;

	_stream["WriteByte"](_type);

	asc_menu_WriteParaBorder(0, _borders.Left, _stream);
	asc_menu_WriteParaBorder(1, _borders.Top, _stream);
	asc_menu_WriteParaBorder(2, _borders.Right, _stream);
	asc_menu_WriteParaBorder(3, _borders.Bottom, _stream);
	asc_menu_WriteParaBorder(4, _borders.InsideH, _stream);
	asc_menu_WriteParaBorder(5, _borders.InsideV, _stream);

	_stream["WriteByte"](255);
}

function asc_menu_ReadCellBackground(_params, _cursor)
{
	var _background = new Asc.CBackground();
	var _continue = true;
	while (_continue)
	{
		var _attr = _params[_cursor.pos++];
		switch (_attr)
		{
			case 0:
			{
				_background.Color = AscCommon.asc_menu_ReadColor(_params, _cursor);
				break;
			}
			case 1:
			{
				_background.Value = _params[_cursor.pos++];
				break;
			}
			case 255:
			default:
			{
				_continue = false;
				break;
			}
		}
	}
	return _background;
}
function asc_menu_WriteCellBackground(_type, _background, _stream)
{
	if (!_background)
		return;

	_stream["WriteByte"](_type);

	asc_menu_WriteColor(0, _background.Color, _stream);

	if (_background.Value !== undefined && _background.Value !== null)
	{
		_stream["WriteByte"](1);
		_stream["WriteLong"](_background.Value);
	}

	_stream["WriteByte"](255);
}

// TABLE
function asc_menu_ReadTableAnchorPosition(_params, _cursor)
{
	var _position = new CTableAnchorPosition();

	_position.CalcX = _params[_cursor.pos++];
	_position.CalcY = _params[_cursor.pos++];
	_position.W = _params[_cursor.pos++];
	_position.H = _params[_cursor.pos++];
	_position.X = _params[_cursor.pos++];
	_position.Y = _params[_cursor.pos++];
	_position.Left_Margin = _params[_cursor.pos++];
	_position.Right_Margin = _params[_cursor.pos++];
	_position.Top_Margin = _params[_cursor.pos++];
	_position.Bottom_Margin = _params[_cursor.pos++];
	_position.Page_W = _params[_cursor.pos++];
	_position.Page_H = _params[_cursor.pos++];
	_position.X_min = _params[_cursor.pos++];
	_position.Y_min = _params[_cursor.pos++];
	_position.X_max = _params[_cursor.pos++];
	_position.Y_max = _params[_cursor.pos++];

	_cursor.pos++;
}
function asc_menu_WriteTableAnchorPosition(_type, _position, _stream)
{
	if (!_position)
		return;

	_stream["WriteByte"](_type);

	_stream["WriteDouble2"](_position.CalcX);
	_stream["WriteDouble2"](_position.CalcY);
	_stream["WriteDouble2"](_position.W);
	_stream["WriteDouble2"](_position.H);
	_stream["WriteDouble2"](_position.X);
	_stream["WriteDouble2"](_position.Y);
	_stream["WriteDouble2"](_position.Left_Margin);
	_stream["WriteDouble2"](_position.Right_Margin);
	_stream["WriteDouble2"](_position.Top_Margin);
	_stream["WriteDouble2"](_position.Bottom_Margin);
	_stream["WriteDouble2"](_position.Page_W);
	_stream["WriteDouble2"](_position.Page_H);
	_stream["WriteDouble2"](_position.X_min);
	_stream["WriteDouble2"](_position.Y_min);
	_stream["WriteDouble2"](_position.X_max);
	_stream["WriteDouble2"](_position.Y_max);

	_stream["WriteByte"](255);
}

function asc_menu_ReadTableLook(_params, _cursor)
{
	var _position = new AscCommon.CTableLook();
	var _continue = true;
	while (_continue)
	{
		var _attr = _params[_cursor.pos++];
		switch (_attr)
		{
			case 0:
			{
				_position.FirstCol = _params[_cursor.pos++];
				break;
			}
			case 1:
			{
				_position.FirstRow = _params[_cursor.pos++];
				break;
			}
			case 2:
			{
				_position.LastCol = _params[_cursor.pos++];
				break;
			}
			case 3:
			{
				_position.LastRow = _params[_cursor.pos++];
				break;
			}
			case 4:
			{
				_position.BandHor = _params[_cursor.pos++];
				break;
			}
			case 5:
			{
				_position.BandVer = _params[_cursor.pos++];
				break;
			}
			case 255:
			default:
			{
				_continue = false;
				break;
			}
		}
	}
	return _position;
}
function asc_menu_WriteTableLook(_type, _look, _stream)
{
	if (!_look)
		return;

	_stream["WriteByte"](_type);

	if (_look.FirstCol !== undefined && _look.FirstCol !== null)
	{
		_stream["WriteByte"](0);
		_stream["WriteBool"](_look.FirstCol);
	}
	if (_look.FirstRow !== undefined && _look.FirstRow !== null)
	{
		_stream["WriteByte"](1);
		_stream["WriteBool"](_look.FirstRow);
	}
	if (_look.LastCol !== undefined && _look.LastCol !== null)
	{
		_stream["WriteByte"](2);
		_stream["WriteBool"](_look.LastCol);
	}
	if (_look.LastRow !== undefined && _look.LastRow !== null)
	{
		_stream["WriteByte"](3);
		_stream["WriteBool"](_look.LastRow);
	}
	if (_look.BandHor !== undefined && _look.BandHor !== null)
	{
		_stream["WriteByte"](4);
		_stream["WriteBool"](_look.BandHor);
	}
	if (_look.BandVer !== undefined && _look.BandVer !== null)
	{
		_stream["WriteByte"](5);
		_stream["WriteBool"](_look.BandVer);
	}

	_stream["WriteByte"](255);
}

function asc_menu_WriteTablePr(_tablePr, _stream)
{
	if (_tablePr.CanBeFlow !== undefined && _tablePr.CanBeFlow !== null)
	{
		_stream["WriteByte"](0);
		_stream["WriteBool"](_tablePr.CanBeFlow);
	}
	if (_tablePr.CellSelect !== undefined && _tablePr.CellSelect !== null)
	{
		_stream["WriteByte"](1);
		_stream["WriteBool"](_tablePr.CellSelect);
	}
	if (_tablePr.TableWidth !== undefined && _tablePr.TableWidth !== null)
	{
		_stream["WriteByte"](2);
		_stream["WriteDouble2"](_tablePr.TableWidth);
	}
	if (_tablePr.TableSpacing !== undefined && _tablePr.TableSpacing !== null)
	{
		_stream["WriteByte"](3);
		_stream["WriteDouble2"](_tablePr.TableSpacing);
	}

	if(_tablePr.TableDefaultMargins) {
		_tablePr.TableDefaultMargins.write(4, _stream);
	}
	asc_menu_WriteCellMargins(5, _tablePr.CellMargins, _stream);

	if (_tablePr.TableAlignment !== undefined && _tablePr.TableAlignment !== null)
	{
		_stream["WriteByte"](6);
		_stream["WriteLong"](_tablePr.TableAlignment);
	}
	if (_tablePr.TableIndent !== undefined && _tablePr.TableIndent !== null)
	{
		_stream["WriteByte"](7);
		_stream["WriteDouble2"](_tablePr.TableIndent);
	}
	if (_tablePr.TableWrappingStyle !== undefined && _tablePr.TableWrappingStyle !== null)
	{
		_stream["WriteByte"](8);
		_stream["WriteLong"](_tablePr.TableWrappingStyle);
	}

	if(_tablePr.TablePaddings) {
		_tablePr.TablePaddings.write(9, _stream);
	}

	asc_menu_WriteCellBorders(10, _tablePr.TableBorders, _stream);
	asc_menu_WriteCellBorders(11, _tablePr.CellBorders, _stream);

	asc_menu_WriteCellBackground(12, _tablePr.TableBackground, _stream);
	asc_menu_WriteCellBackground(13, _tablePr.CellsBackground, _stream);

	asc_menu_WritePosition(14, _tablePr.Position, _stream);
	asc_menu_WriteImagePosition(15, _tablePr.PositionH, _stream);
	asc_menu_WriteImagePosition(16, _tablePr.PositionV, _stream);

	asc_menu_WriteTableAnchorPosition(17, _tablePr.Internal_Position, _stream);

	if (_tablePr.ForSelectedCells !== undefined && _tablePr.ForSelectedCells !== null)
	{
		_stream["WriteByte"](18);
		_stream["WriteBool"](_tablePr.ForSelectedCells);
	}
	if (_tablePr.TableStyle !== undefined && _tablePr.TableStyle !== null)
	{
		_stream["WriteByte"](19);
		_stream["WriteString2"](_tablePr.TableStyle);
	}

	asc_menu_WriteTableLook(20, _tablePr.TableLook, _stream);

	if (_tablePr.RowsInHeader !== undefined && _tablePr.RowsInHeader !== null)
	{
		_stream["WriteByte"](21);
		_stream["WriteLong"](_tablePr.RowsInHeader);
	}
	if (_tablePr.CellsVAlign !== undefined && _tablePr.CellsVAlign !== null)
	{
		_stream["WriteByte"](22);
		_stream["WriteLong"](_tablePr.CellsVAlign);
	}
	if (_tablePr.AllowOverlap !== undefined && _tablePr.AllowOverlap !== null)
	{
		_stream["WriteByte"](23);
		_stream["WriteBool"](_tablePr.AllowOverlap);
	}
	if (_tablePr.TableLayout !== undefined && _tablePr.TableLayout !== null)
	{
		_stream["WriteByte"](24);
		_stream["WriteLong"](_tablePr.TableLayout);
	}
	if (_tablePr.Locked !== undefined && _tablePr.Locked !== null)
	{
		_stream["WriteByte"](25);
		_stream["WriteBool"](_tablePr.Locked);
	}

	_stream["WriteByte"](255);
}

// HYPERLINKS
function asc_menu_ReadHyperPr(_params, _cursor)
{
	var _settings = new Asc.CHyperlinkProperty();

	var _continue = true;
	while (_continue)
	{
		var _attr = _params[_cursor.pos++];
		switch (_attr)
		{
			case 0:
			{
				_settings.Text = _params[_cursor.pos++];
				break;
			}
			case 1:
			{
				_settings.Value = _params[_cursor.pos++];
				break;
			}
			case 2:
			{
				_settings.ToolTip = _params[_cursor.pos++];
				break;
			}
			case 255:
			default:
			{
				_continue = false;
				break;
			}
		}
	}

	return _settings;
}
function asc_menu_WriteHyperPr(_hyperPr, _stream)
{
	if (_hyperPr.Text !== undefined && _hyperPr.Text !== null)
	{
		_stream["WriteByte"](0);
		_stream["WriteString2"](_hyperPr.Text);
	}

	if (_hyperPr.Value !== undefined && _hyperPr.Value !== null)
	{
		_stream["WriteByte"](1);
		_stream["WriteString2"](_hyperPr.Value);
	}

	if (_hyperPr.ToolTip !== undefined && _hyperPr.ToolTip !== null)
	{
		_stream["WriteByte"](2);
		_stream["WriteString2"](_hyperPr.ToolTip);
	}

	_stream["WriteByte"](255);
}

// COLORSCHEME
function asc_WriteColorSchemes(schemas, s)
{
	s["WriteLong"](schemas.length);

	for (var i = 0; i < schemas.length; ++i)
	{
		s["WriteString2"](schemas[i].get_name());

		var colors = schemas[i].get_colors();
		s["WriteLong"](colors.length);

		for (var j = 0; j < colors.length; ++j)
		{
			asc_menu_WriteColor(0, colors[j], s);
		}
	}
}

var global_memory_stream_menu = CreateEmbedObject("CMemoryStreamEmbed");

// SERIALIZE EVENTS

// COMMON
function stringOOToLocalDate(date)
{
	if (typeof date === 'string')
		return parseInt(date);
	return 0;
}

function stringUtcToLocalDate(date)
{
	if (typeof date === 'string')
		return parseInt(date) + (new Date()).getTimezoneOffset() * 60000;
	return 0;
}

function getHexColor(r, g, b)
{
	r = r.toString(16);
	g = g.toString(16);
	b = b.toString(16);
	if (r.length === 1) r = '0' + r;
	if (g.length === 1) g = '0' + g;
	if (b.length === 1) b = '0' + b;
	return r + g + b;
}

function postDataAsJSONString(data, eventId)
{
	var stream = global_memory_stream_menu;
	stream["ClearNoAttack"]();
	if (data !== undefined && data !== null) {
		stream["WriteString2"](JSON.stringify(data));
	}
	window["native"]["OnCallMenuEvent"](eventId, stream);
}

// COMMENTS
function readSDKComment(id, data)
{
	var date = data.asc_getOnlyOfficeTime()
			? new Date(stringOOToLocalDate(data.asc_getOnlyOfficeTime()))
			: (data.asc_getTime() == '') ? new Date() : new Date(stringUtcToLocalDate(data.asc_getTime())),
		groupname = id.substr(0, id.lastIndexOf('_') + 1).match(/^(doc|sheet[0-9_]+)_/);

	return {
		id          : id,
		guid        : data.asc_getGuid(),
		userId      : data.asc_getUserId(),
		userName    : data.asc_getUserName(),
		date        : date.getTime().toString(),
		quoteText   : data.asc_getQuoteText(),
		text        : data.asc_getText(),
		solved      : data.asc_getSolved(),
		unattached  : (data.asc_getDocumentFlag === undefined) ? false : data.asc_getDocumentFlag(),
		groupName   : (groupname && groupname.length>1) ? groupname[1] : null,
		replies     : readSDKReplies(data)
	};
}

function readSDKReplies(data)
{
	var i = 0,
		replies = [],
		date = null;
	var repliesCount = data.asc_getRepliesCount();
	if (repliesCount)
	{
		for (i = 0; i < repliesCount; ++i)
		{
			var reply = data.asc_getReply(i);
			date = (reply.asc_getOnlyOfficeTime())
				? new Date(stringOOToLocalDate(reply.asc_getOnlyOfficeTime()))
				: ((reply.asc_getTime() == '') ? new Date() : new Date(stringUtcToLocalDate(reply.asc_getTime())));
			replies.push({
				userId      : reply.asc_getUserId(),
				userName    : reply.asc_getUserName(),
				text        : reply.asc_getText(),
				date        : date.getTime().toString()
			});
		}
	}
	return replies;
}

function onApiAddComment(id, data)
{
	function _apply() {
		var comment = readSDKComment(id, data) || {};
		postDataAsJSONString(comment, 23001); // ASC_MENU_EVENT_TYPE_ADD_COMMENT
	}
	if (AscCommon.isApiAddCommentAsync === true)
		setTimeout(_apply, 5);
	else
		_apply();
}

function onApiAddComments(data)
{
	function _apply() {
		var comments = [];
		for (var i = 0; i < data.length; ++i)
			comments.push(readSDKComment(data[i].asc_getId(), data[i]));
		postDataAsJSONString(comments, 23002); // ASC_MENU_EVENT_TYPE_ADD_COMMENTS
	}
	if (AscCommon.isApiAddCommentAsync === true)
		setTimeout(_apply, 5);
	else
		_apply();
}

function onApiRemoveComment(id)
{
	var data = {
		"id": id
	};
	postDataAsJSONString(data, 23003); // ASC_MENU_EVENT_TYPE_REMOVE_COMMENT
}

function onApiChangeComments(data)
{
	var comments = [];
	for (var i = 0; i < data.length; ++i)
		comments.push(readSDKComment(data[i].asc_getId(), data[i]));
	postDataAsJSONString(comments, 23004); // ASC_MENU_EVENT_TYPE_CHANGE_COMMENTS
}

function onApiRemoveComments(data)
{
	var ids = [];
	for (var i = 0; i < data.length; ++i)
	{
		ids.push({
			"id": data[i]
		});
	}
	postDataAsJSONString(ids, 23005); // ASC_MENU_EVENT_TYPE_REMOVE_COMMENTS
}

function onApiChangeCommentData(id, data)
{
	var comment = readSDKComment(id, data) || {},
		change = {
			"id": id,
			"comment": comment
		};

	postDataAsJSONString(change, 23006); // ASC_MENU_EVENT_TYPE_CHANGE_COMMENTDATA
}

function onApiLockComment(id, userId)
{
	var data = {
		"id": id,
		"userId": userId
	};
	postDataAsJSONString(data, 23007); // ASC_MENU_EVENT_TYPE_LOCK_COMMENT
}

function onApiUnLockComment(id)
{
	var data = {
		"id": id
	};
	postDataAsJSONString(data, 23008); // ASC_MENU_EVENT_TYPE_UNLOCK_COMMENT
}

function onApiShowComment(uids, posX, posY, leftX, opts, hint)
{
	var data = {
		"uids": uids,
		"posX": posX,
		"posY": posY,
		"leftX": leftX,
		"opts": opts,
		"hint": hint
	};
	postDataAsJSONString(data, 23009); // ASC_MENU_EVENT_TYPE_SHOW_COMMENT
}

function onApiHideComment(hint)
{
	var data = {
		"hint": hint
	};
	postDataAsJSONString(data, 23010); // ASC_MENU_EVENT_TYPE_HIDE_COMMENT
}

function onApiUpdateCommentPosition(uids, posX, posY, leftX)
{
	var data = {
		"uids": uids,
		"posX": posX,
		"posY": posY,
		"leftX": leftX
	};
	postDataAsJSONString(data, 23011); // ASC_MENU_EVENT_TYPE_UPDATE_COMMENT_POSITION
}

// USERS
function sdkUsersToJson(users)
{
	var arrUsers = [];

	for (var userId in users)
	{
		if (undefined !== userId)
		{
			var user = users[userId];
			if (user) {
				arrUsers.push({
					id          : user.asc_getId(),
					idOriginal  : user.asc_getIdOriginal(),
					userName    : user.asc_getUserName(),
					online      : true,
					color       : user.asc_getColor(),
					view        : user.asc_getView()
				});
			}
		}
	}
	return arrUsers;
}

function onApiAuthParticipantsChanged(users)
{
	var _users = sdkUsersToJson(users) || [];
	postDataAsJSONString(_users, 20101); // ASC_COAUTH_EVENT_TYPE_PARTICIPANTS_CHANGED
}

function onApiParticipantsChanged(users)
{
	var _users = sdkUsersToJson(users) || [];
	postDataAsJSONString(_users, 20101); // ASC_COAUTH_EVENT_TYPE_PARTICIPANTS_CHANGED
}

// PLACE
function onDocumentPlaceChanged()
{
	postDataAsJSONString(null, 23012); // ASC_MENU_EVENT_TYPE_DOCUMENT_PLACE_CHANGED
}

// ACTIONS
function onApiLongActionBegin(type, id)
{
	var info = {
		"type" : type,
		"id" : id
	};
	postDataAsJSONString(info, 26102); // ASC_MENU_EVENT_TYPE_LONGACTION_BEGIN
}

function onApiLongActionEnd(type, id)
{
	var info = {
		"type" : type,
		"id" : id
	};
	postDataAsJSONString(info, 26103); // ASC_MENU_EVENT_TYPE_LONGACTION_END
}

// ERRORS
function onApiError(id, level, errData)
{
	var info = {
		"level" : level,
		"id" : id,
		"errData" : JSON.prune(errData, 4)
	};
	postDataAsJSONString(info, 26104); // ASC_MENU_EVENT_TYPE_API_ERROR
}

// THEME COLORS
function onApiSendThemeColors(theme_colors, standart_colors)
{
	var colors = {
		"themeColors": theme_colors.map(function(color) {
			return getHexColor(color.get_r(), color.get_g(), color.get_b());
		})
	}
	if (standart_colors != null) {
		colors["standartColors"] = standart_colors.map(function(color) {
			return getHexColor(color.get_r(), color.get_g(), color.get_b());
		});
	}
	postDataAsJSONString(colors, 2417); // ASC_MENU_EVENT_TYPE_THEMECOLORS
}

// JSON.prune
// JSON.prune : a function to stringify any object without overflow
// two additional optional parameters :
//   - the maximal depth (default : 6)
//   - the maximal length of arrays (default : 50)
// You can also pass an "options" object.
// examples :
//   var json = JSON.prune(window)
//   var arr = Array.apply(0,Array(1000)); var json = JSON.prune(arr, 4, 20)
//   var json = JSON.prune(window.location, {inheritedProperties:true})
// Web site : http://dystroy.org/JSON.prune/
// JSON.prune on github : https://github.com/Canop/JSON.prune
// This was discussed here : http://stackoverflow.com/q/13861254/263525
// The code is based on Douglas Crockford's code : https://github.com/douglascrockford/JSON-js/blob/master/json2.js
// No effort was done to support old browsers. JSON.prune will fail on IE8.
(function () {
	'use strict';

	var DEFAULT_MAX_DEPTH = 6;
	var DEFAULT_ARRAY_MAX_LENGTH = 50;
	var DEFAULT_PRUNED_VALUE = '"-pruned-"';
	var seen; // Same variable used for all stringifications
	var iterator; // either forEachEnumerableOwnProperty, forEachEnumerableProperty or forEachProperty

	// iterates on enumerable own properties (default behavior)
	var forEachEnumerableOwnProperty = function(obj, callback) {
		for (var k in obj) {
			if (Object.prototype.hasOwnProperty.call(obj, k)) callback(k);
		}
	};
	// iterates on enumerable properties
	var forEachEnumerableProperty = function(obj, callback) {
		for (var k in obj) callback(k);
	};
	// iterates on properties, even non enumerable and inherited ones
	// This is dangerous
	var forEachProperty = function(obj, callback, excluded) {
		if (obj==null) return;
		excluded = excluded || {};
		Object.getOwnPropertyNames(obj).forEach(function(k){
			if (!excluded[k]) {
				callback(k);
				excluded[k] = true;
			}
		});
		forEachProperty(Object.getPrototypeOf(obj), callback, excluded);
	};

	Object.defineProperty(Date.prototype, "toPrunedJSON", {value:Date.prototype.toJSON});

	var	cx = /[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,
		escapable = /[\\\"\x00-\x1f\x7f-\x9f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,
		meta = {	// table of character substitutions
			'\b': '\\b',
			'\t': '\\t',
			'\n': '\\n',
			'\f': '\\f',
			'\r': '\\r',
			'"' : '\\"',
			'\\': '\\\\'
		};

	function quote(string) {
		escapable.lastIndex = 0;
		return escapable.test(string) ? '"' + string.replace(escapable, function (a) {
			var c = meta[a];
			return typeof c === 'string'
				? c
				: '\\u' + ('0000' + a.charCodeAt(0).toString(16)).slice(-4);
		}) + '"' : '"' + string + '"';
	}

	var prune = function (value, depthDecr, arrayMaxLength) {
		var prunedString = DEFAULT_PRUNED_VALUE;
		var replacer;
		if (typeof depthDecr == "object") {
			var options = depthDecr;
			depthDecr = options.depthDecr;
			arrayMaxLength = options.arrayMaxLength;
			iterator = options.iterator || forEachEnumerableOwnProperty;
			if (options.allProperties) iterator = forEachProperty;
			else if (options.inheritedProperties) iterator = forEachEnumerableProperty
			if ("prunedString" in options) {
				prunedString = options.prunedString;
			}
			if (options.replacer) {
				replacer = options.replacer;
			}
		} else {
			iterator = forEachEnumerableOwnProperty;
		}
		seen = [];
		depthDecr = depthDecr || DEFAULT_MAX_DEPTH;
		arrayMaxLength = arrayMaxLength || DEFAULT_ARRAY_MAX_LENGTH;
		function str(key, holder, depthDecr) {
			var i, k, v, length, partial, value = holder[key];

			if (value && typeof value === 'object' && typeof value.toPrunedJSON === 'function') {
				value = value.toPrunedJSON(key);
			}
			if (value && typeof value.toJSON === 'function') {
				value = value.toJSON();
			}

			switch (typeof value) {
				case 'string':
					return quote(value);
				case 'number':
					return isFinite(value) ? String(value) : 'null';
				case 'boolean':
				case 'null':
					return String(value);
				case 'object':
					if (!value) {
						return 'null';
					}
					if (depthDecr<=0 || seen.indexOf(value)!==-1) {
						if (replacer) {
							var replacement = replacer(value, prunedString, true);
							return replacement===undefined ? undefined : ''+replacement;
						}
						return prunedString;
					}
					seen.push(value);
					partial = [];
					if (Object.prototype.toString.apply(value) === '[object Array]') {
						length = Math.min(value.length, arrayMaxLength);
						for (i = 0; i < length; i += 1) {
							partial[i] = str(i, value, depthDecr-1) || 'null';
						}
						v = '[' + partial.join(',') + ']';
						if (replacer && value.length>arrayMaxLength) return replacer(value, v, false);
						return v;
					}
					if (value instanceof RegExp) {
						return quote(value.toString());
					}
					iterator(value, function(k) {
						try {
							v = str(k, value, depthDecr-1);
							if (v) partial.push(quote(k) + ':' + v);
						} catch (e) {
							// this try/catch due to forbidden accessors on some objects
						}
					});
					return '{' + partial.join(',') + '}';
				case 'function':
				case 'undefined':
					return replacer ? replacer(value, undefined, false) : undefined;
			}
		}
		return str('', {'': value}, depthDecr);
	};

	prune.log = function() {
		console.log.apply(console, Array.prototype.map.call(arguments, function(v) {
			return JSON.parse(JSON.prune(v));
		}));
	};
	prune.forEachProperty = forEachProperty; // you might want to also assign it to Object.forEachProperty

	if (typeof module !== "undefined") module.exports = prune;
	else JSON.prune = prune;
}());
