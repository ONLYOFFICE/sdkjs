(function () {
	function PropLocker(objectId) {
		this.objectId = null;
		this.Lock = new AscCommon.CLock();
		this.Id = AscCommon.g_oIdCounter.Get_NewId();
		g_oTableId.Add(this, this.Id);

		if (typeof objectId === "string") {
			this.setObjectId(objectId);
		}

	}

	PropLocker.prototype.getObjectType = function () {
		return AscDFH.historyitem_type_PropLocker;
	};
	PropLocker.prototype.setObjectId = function (id) {
		History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_PropLockerSetId, this.objectId, id));
		this.objectId = id;
	};
	PropLocker.prototype.Get_Id = function () {
		return this.Id;
	};
	PropLocker.prototype.Write_ToBinary2 = function (w) {
		w.WriteLong(AscDFH.historyitem_type_PropLocker);
		w.WriteString2(this.Id);
	};

	PropLocker.prototype.Read_FromBinary2 = function (r) {
		this.Id = r.GetString2();
	};

	PropLocker.prototype.Refresh_RecalcData = function () {};

	window["AscCommon"] = window["AscCommon"] || {};
	window["AscCommon"].PropLocker = PropLocker;
})();
