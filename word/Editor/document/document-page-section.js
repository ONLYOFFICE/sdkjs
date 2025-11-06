/*
 * (c) Copyright Ascensio System SIA 2010-2025
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


(function()
{
	/**
	 * @constructor
	 */
	function DocumentPageSection()
	{
		this.Pos    = 0;
		this.EndPos = -1;
		
		this.ResetSectionStart = false; // Если секция разбилась внутри блочного контент контрола
		
		this.Y      = 0;
		this.YLimit = 0;
		
		this.YLimit2 = 0;
		
		this.Index   = -1;
		this.SectPr  = null;
		
		this.Columns = [];
		this.ColumnsSep = false;
		
		this.IterationsCount       = 0;
		this.CurrentY              = 0;
		this.RecalculateBottomLine = true;
		this.CanDecrease           = true;
		this.WasIncrease           = false; // Было ли хоть раз увеличение
		this.IterationStep         = 10;
		this.IterationDirection    = 0;
	}
	/**
	 * Инициализируем параметры данной секции
	 * @param {number} nPageAbs
	 * @param {AscWord.SectPr} oSectPr
	 * @param {Number} nSectionIndex
	 */
	DocumentPageSection.prototype.Init = function(nPageAbs, oSectPr, nSectionIndex)
	{
		var oFrame  = oSectPr.GetContentFrame(nPageAbs);
		var nX      = oFrame.Left;
		var nXLimit = oFrame.Right;
		
		for (var nCurColumn = 0, nColumnsCount = oSectPr.GetColumnCount(); nCurColumn < nColumnsCount; ++nCurColumn)
		{
			this.Columns[nCurColumn] = new AscWord.DocumentPageColumn();
			
			this.Columns[nCurColumn].X      = nX;
			this.Columns[nCurColumn].XLimit = nColumnsCount - 1 === nCurColumn ? nXLimit : nX + oSectPr.GetColumnWidth(nCurColumn);
			
			nX += oSectPr.GetColumnWidth(nCurColumn) + oSectPr.GetColumnSpace(nCurColumn);
		}
		this.ColumnsSep = oSectPr.GetColumnSep();
		
		this.Y       = oFrame.Top;
		this.YLimit  = oFrame.Bottom;
		this.YLimit2 = oFrame.Bottom;
		this.Index   = nSectionIndex;
		this.SectPr  = oSectPr;
	};
	DocumentPageSection.prototype.GetIndex = function()
	{
		return this.Index;
	};
	DocumentPageSection.prototype.GetSectPr = function()
	{
		return this.SectPr;
	};
	DocumentPageSection.prototype.Copy = function()
	{
		var NewSection = new DocumentPageSection();
		
		NewSection.Pos    = this.Pos;
		NewSection.EndPos = this.EndPos;
		NewSection.Y      = this.Y;
		NewSection.YLimit = this.YLimit;
		
		for (var ColumnIndex = 0, Count = this.Columns.length; ColumnIndex < Count; ++ColumnIndex)
		{
			NewSection.Columns[ColumnIndex] = this.Columns[ColumnIndex].Copy();
		}
		
		return NewSection;
	};
	DocumentPageSection.prototype.Shift = function(Dx, Dy)
	{
		this.Y      += Dy;
		this.YLimit += Dy;
		
		for (var ColumnIndex = 0, Count = this.Columns.length; ColumnIndex < Count; ++ColumnIndex)
		{
			this.Columns[ColumnIndex].Shift(Dx, Dy);
		}
	};
	DocumentPageSection.prototype.CorrectY = function(y)
	{
		return (y < this.Y ? this.Y : (y > this.YLimit ? this.YLimit : y));
	};
	/**
	 * Происходи ли процесс расчета нижней границы разрыва секции на текущей странице
	 * @returns {boolean}
	 */
	DocumentPageSection.prototype.IsCalculatingSectionBottomLine = function()
	{
		return (this.IterationsCount > 0 && true === this.RecalculateBottomLine);
	};
	/**
	 * Можно ли расчитывать нижнюю границу разрыва секции на текущей странице
	 * @returns {boolean}
	 */
	DocumentPageSection.prototype.CanRecalculateBottomLine = function()
	{
		return this.RecalculateBottomLine;
	};
	/**
	 * Запрещаем возможность расчета нижней границы разрыва секции на текущей страницы
	 */
	DocumentPageSection.prototype.ForbidRecalculateBottomLine = function()
	{
		this.RecalculateBottomLine = false;
	};
	DocumentPageSection.prototype.GetY = function()
	{
		return this.Y;
	};
	DocumentPageSection.prototype.GetYLimit = function()
	{
		if (0 === this.IterationsCount)
			return this.YLimit;
		else
			return this.CurrentY;
	};
	/**
	 * Производим шаг рассчета нижней границы рарзрыва секции
	 * @param {boolean} isIncrease
	 * @returns {number}
	 */
	DocumentPageSection.prototype.IterateBottomLineCalculation = function(isIncrease)
	{
		// Алгоритм следующий:
		// На первом шаге мы прогнозируем положение границы по уже имеющемуся объему текста и
		// ширине колонок.
		// Далее мы сдвигаем границу на значение IterationStep, вверх или вниз в зависимости
		// от результата расчета страницы. При перемене направления сдвига мы всегда уменьшаем шаг
		// в 2 раза. Также мы можем уменьшать шаг в 2 раза, если сдвигаем границу вверх, и хотябы
		// раз до этого двигали ее вниз. Останавливаем итерацию при попытке подвинуть границу наверх,
		// когда шаг итерации становится менее 2мм.
		
		if (0 === this.IterationsCount)
		{
			// Пытаемся заранее спрогнозировать позицию, где должно быть разделение. Учитывая, что колонки могут быть разной
			// ширины, мы расчитываем суммарную занимаемую текстом область. Делим ее по колонкам, с учетом их суммарной
			// ширины.
			
			var nSumArea = 0, nSumWidth = 0;
			for (var nColumnIndex = 0, nColumnsCount = this.Columns.length; nColumnIndex < nColumnsCount; ++nColumnIndex)
			{
				var oColumn = this.Columns[nColumnIndex];
				if (true !== oColumn.Empty)
					nSumArea += (oColumn.Bounds.Bottom - this.Y) * (oColumn.XLimit - oColumn.X);
				
				nSumWidth += oColumn.XLimit - oColumn.X;
			}
			
			if (nSumWidth > 0.001)
				this.CurrentY = this.Y + nSumArea / nSumWidth;
			else
				this.CurrentY = this.Y;
		}
		else
		{
			if (false === isIncrease)
			{
				if (this.IterationDirection > 0 || this.WasIncrease)
					this.IterationStep /= 2;
				
				this.CurrentY -= this.IterationStep;
				this.IterationDirection = -1;
				
				if (this.CurrentY < this.Y)
				{
					// Такое может быть, когда у нас всего одна строка в начале страницы, которую мы размещаем всегда
					this.CurrentY    = this.Y;
					this.CanDecrease = false;
				}
			}
			else
			{
				if (this.IterationDirection < 0)
					this.IterationStep /= 2;
				
				this.CurrentY += this.IterationStep;
				this.IterationDirection = 1;
				
				this.WasIncrease = true;
			}
		}
		
		if (this.IterationStep < 2)
			this.CanDecrease = false;
		
		this.CurrentY = Math.min(this.CurrentY, this.YLimit2);
		
		// // Recalculation LOG
		// console.log(`Calculate continuous section count=${this.IterationsCount} step=${this.IterationStep} Y=${this.CurrentY}`);
		
		this.IterationsCount++;
		return this.CurrentY;
	};
	DocumentPageSection.prototype.Reset_Columns = function()
	{
		for (var ColumnIndex = 0, Count = this.Columns.length; ColumnIndex < Count; ++ColumnIndex)
		{
			this.Columns[ColumnIndex].Reset();
		}
	};
	/**
	 * Можем ли мы провести еще одну итерацию с уменьшением нижней границы
	 * @returns {boolean}
	 */
	DocumentPageSection.prototype.CanDecreaseBottomLine = function()
	{
		return this.CanDecrease;
	};
	DocumentPageSection.prototype.CanIncreaseBottomLine = function()
	{
		// Данная функция не должна возвращать false, если возвращает, значит неправильно работает алгоритм по вычислению
		// нижней границы continuous секции
		// if (!(this.YLimit2 - this.CurrentY > 0.001))
		// 	console.log("Bad continuous section calculate");
		
		return (this.YLimit2 - this.CurrentY > 0.001);
	};
	DocumentPageSection.prototype.GetBottomLimit = function()
	{
		let bottomLimit = this.Y;
		for (let column = 0, columnCount = this.Columns.length; column < columnCount; ++column)
		{
			bottomLimit = Math.max(bottomLimit, this.Columns[column].Bounds.Bottom);
		}
		return bottomLimit;
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.DocumentPageSection = DocumentPageSection;
	
})();
