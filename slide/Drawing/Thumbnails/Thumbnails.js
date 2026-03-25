function CThPosBase(oThumbnails)
{
	this.thumbnails = oThumbnails;
}
CThPosBase.prototype.Increment = function() {};
CThPosBase.prototype.Decrement = function() {};
CThPosBase.prototype.IsLess = function(oPos)
{
	return false;
};
CThPosBase.prototype.IsGreater = function(oPos)
{
	return false;
};
CThPosBase.prototype.IsLessOrEqual = function(oPos)
{
	return this.IsLess(oPos) || this.IsEqual(oPos);
};
CThPosBase.prototype.IsGreaterOrEqual = function(oPos)
{
	return this.IsGreater(oPos) || this.IsEqual(oPos);
};
CThPosBase.prototype.IsEqual = function(oPos)
{
	return false;
};
CThPosBase.prototype.Copy = function()
{
	return new CThPosBase(this.thumbnails);
};
CThPosBase.prototype.GetPage = function()
{
	return this.thumbnails.GetPage(this);
};
CThPosBase.prototype.GetSld = function()
{
	return null;
};
CThPosBase.prototype.GetPresentation = function()
{
	return Asc.editor.WordControl.m_oLogicDocument;
};

function CSlideThPos(oThumbnails)
{
	CThPosBase.call(this, oThumbnails);
	this.Idx = -1;
}
AscFormat.InitClassWithoutType(CSlideThPos, CThPosBase);
CSlideThPos.prototype.Check = function()
{
	let aThs = this.thumbnails.slides;
	let nMaxIdx = aThs.length - 1;
	let nMinIdx = Math.min(nMaxIdx, 0);
	this.Idx = Math.max(nMinIdx, Math.min(this.Idx, nMaxIdx));
};
CSlideThPos.prototype.Increment = function()
{
	++this.Idx;
	this.Check();
};
CSlideThPos.prototype.Decrement = function()
{
	--this.Idx;
	this.Check();
};
CSlideThPos.prototype.IsLess = function(oPos)
{
	return this.Idx < oPos.Idx;
};
CSlideThPos.prototype.IsGreater = function(oPos)
{
	return this.Idx > oPos.Idx;
};
CSlideThPos.prototype.IsEqual = function(oPos)
{
	return this.Idx === oPos.Idx;
};
CSlideThPos.prototype.Copy = function()
{
	let oPos = new CSlideThPos(this.thumbnails);
	oPos.Idx = this.Idx;
	return oPos;
};
CSlideThPos.prototype.GetSld = function()
{
	return this.GetPresentation().Slides[this.Idx];
};

function CThumbnailsBase() {
	this.curPos = this.GetIterator();
	this.firstVisible = this.GetIterator();
	this.lastVisible = this.GetIterator();
}
CThumbnailsBase.prototype.GetIterator = function()
{
	throw new Error();
};
CThumbnailsBase.prototype.SetSlideRecalc = function (idx) {
	let oTh = this.GetPage(idx);
	if(oTh)
	{
		oTh.SetRecalc(true);
	}
};
CThumbnailsBase.prototype.GetPage = function (oIdx) {
	return null;
};
CThumbnailsBase.prototype.SelectAll = function()
{
	this.IteratePages(this.GetStartPos(), this.GetEndPos(), function(oPage) {
		oPage.SetSelected(true);
		return false;
	});
};
CThumbnailsBase.prototype.FindThUnderCursor = function(X, Y)
{
	this.IteratePages(this.GetStartPos(), this.GetEndPos(), function(oPage) {
		if(oPage.Hit(X, Y))
		{
			return oPage;
		}
	});
};
CThumbnailsBase.prototype.ResetSelection = function()
{
	this.IteratePages(this.GetStartPos(), this.GetEndPos(), function(oPage) {
		oPage.SetSelected(false);
		return false;
	});
};
CThumbnailsBase.prototype.SetAllFocused = function(bValue)
{
	this.IteratePages(this.GetStartPos(), this.GetEndPos(), function(oPage) {
		oPage.SetFocused(bValue);
		return false;
	});
};

CThumbnailsBase.prototype.SetAllRecalc = function (bValue) {
	this.IteratePages(this.GetStartPos(), this.GetEndPos(), function(oPage) {
		oPage.SetRecalc(bValue);
		return false;
	});
};
CThumbnailsBase.prototype.SelectRange = function(oStartPos, oEndPos)
{
	this.SetSelectionRange(oStartPos, oEndPos, true);
};

CThumbnailsBase.prototype.SetSelectionRange = function(oStartPos, oEndPos, bValue)
{
	this.IteratePages(oStartPos, oEndPos, function(oPage) {
		oPage.SetSelected(bValue);
		return false;
	});
};
CThumbnailsBase.prototype.GetStartPos = function()
{
	throw new Error();
};
CThumbnailsBase.prototype.GetEndPos = function()
{
	throw new Error();
};
CThumbnailsBase.prototype.IteratePages = function(oStartPos, oEndPos, fAction)
{
	if(oStartPos.IsLessOrEqual(oEndPos))
	{
		for(let oPos = oStartPos.Copy(); oPos.IsLessOrEqual(oEndPos); oPos.Increment())
		{
			let oPage = oPos.GetPage();
			if(!oPage)
				return null;
			if(oPage)
			{
				let oResult = fAction(oPos.GetPage());
				if(oResult)
				{
					return oResult;
				}
			}
		}
	}
	return null;
};
CThumbnailsBase.prototype.GetPagesCount = function()
{
	let nCount = 0;
	this.IteratePages(this.GetStartPos(), this.GetEndPos(), function(oPage) {
		nCount++;
		return false;
	});
	return nCount;
};
CThumbnailsBase.prototype.ConvertCoords2 = function(X, Y)
{
	let pages_count = this.GetPagesCount();
	if (0 == pages_count)
		return -1;

	let oStartPos = this.GetStartPos();
	let oEndPos = this.GetEndPos();
	let _min = Math.abs(Y - this.GetStartPos().GetPage().top);
	let _MinPositionPage = 0;

	this.IteratePages(oStartPos, oEndPos, function(oPage) {

		let _min1 = Math.abs(Y - oPage.top);
		let _min2 = Math.abs(Y - oPage.bottom);

		if (_min1 < _min)
		{
			_min = _min1;
			_MinPositionPage = oPage.GetPosition();
		}
		if (_min2 < _min)
		{
			_min = _min2;
			_MinPositionPage = oPage.GetPosition().Increment();
		}
	});

	return _MinPositionPage;
};
CThumbnailsBase.prototype.GetPagePosition = function(oPageIdx)
{
	let oPage = this.GetPage(oPageIdx);
	if(!oPage) return null;
	let _ret = {
		X: AscCommon.AscBrowser.convertToRetinaValue(oPage.left),
		Y: AscCommon.AscBrowser.convertToRetinaValue(oPage.top),
		W: AscCommon.AscBrowser.convertToRetinaValue(oPage.right - oPage.left),
		H: AscCommon.AscBrowser.convertToRetinaValue(oPage.bottom - oPage.top)
	};
	return _ret;
};
CThumbnailsBase.prototype.GetCurSld = function()
{
	return this.curPos.GetSld();
};
CThumbnailsBase.prototype.IndexToPos = function(nIndex)
{
	return null;
};
CThumbnailsBase.prototype.PosToIndex = function(oPos)
{
	return null;
};

function CSlidesThumbnails() {
	CThumbnailsBase.call(this);
	this.slides = [];
}
AscFormat.InitClassWithoutType(CSlidesThumbnails, CThumbnailsBase);
CSlidesThumbnails.prototype.GetIterator = function()
{
	return new CSlideThPos(this);
};
CSlidesThumbnails.prototype.clear = function ()
{
	this.slides.length = 0;
	this.cusSlide = -1;
	this.firstVisible = -1;
	this.lastVisible = -1;
};
CSlidesThumbnails.prototype.GetStartPos = function()
{
	let oPos = this.GetIterator();
	oPos.Idx = Math.min(0, this.slides.length - 1);
	return oPos;
};
CSlidesThumbnails.prototype.GetEndPos = function()
{
	let oPos = this.GetIterator();
	oPos.Idx = this.slides.length - 1;
	return oPos;
};
CSlidesThumbnails.prototype.IndexToPos = function(nIndex)
{
	let oPos = this.GetIterator();
	oPos.Idx = nIndex;
	return oPos;
};
CSlidesThumbnails.prototype.PosToIndex = function(oPos)
{
	if(!oPos) return -1;
	return oPos.Idx;
};
CSlidesThumbnails.prototype.PosToIndex = function(oPos)
{
	let nIdx = -1;
	this.IteratePages(this.GetStartPos(), this.GetEndPos(), function(oPage, oPosition) {
		++nIdx;
		if(oPos.IsEqual(oPosition)) {
			return true;
		}
		return false;
	});
	return nIdx;
};

window["AscCommon"] = window["AscCommon"] || {};
window["AscCommon"].CSlidesThumbnails = CSlidesThumbnails;