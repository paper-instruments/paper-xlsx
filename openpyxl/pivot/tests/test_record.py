# Copyright (c) 2010-2024 openpyxl
import pytest

from io import BytesIO
from zipfile import ZipFile

from openpyxl.packaging.manifest import Manifest
from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from .test_fields import (
    Index,
    Number,
    Text,
)

@pytest.fixture
def Record():
    from ..record import Record
    return Record


class TestRecord:

    def test_ctor(self, Record, Number, Text, Index):
        n = [Number(v=1), Number(v=25)]
        s = [Text(v="2014-03-24")]
        x = [Index(), Index(), Index()]
        fields = n + s + x
        field = Record(_fields=fields)
        xml = tostring(field.to_tree())
        expected = """
        <r>
          <n v="1"/>
          <n v="25"/>
          <s v="2014-03-24"/>
          <x v="0"/>
          <x v="0"/>
          <x v="0"/>
        </r>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Record, Number, Text, Index):
        src = """
        <r>
          <n v="1"/>
          <x v="0"/>
          <s v="2014-03-24"/>
          <x v="0"/>
          <n v="25"/>
          <x v="0"/>
        </r>
        """
        node = fromstring(src)
        n = [Number(v=1), Number(v=25)]
        s = [Text(v="2014-03-24")]
        x = [Index(), Index(), Index()]
        fields = [
            Number(v=1),
            Index(),
            Text(v="2014-03-24"),
            Index(),
            Number(v=25),
            Index(),
        ]
        field = Record.from_tree(node)
        assert field == Record(_fields=fields)


@pytest.fixture
def RecordList():
    from ..record import RecordList
    return RecordList


class TestRecordList:

    def test_ctor(self, RecordList):
        cache = RecordList()
        xml = tostring(cache.to_tree())
        expected = """
        <pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           count="0" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, RecordList):
        src = """
        <pivotCacheRecords count="0" />
        """
        node = fromstring(src)
        cache = RecordList.from_tree(node)
        assert cache == RecordList()


    def test_write(self, RecordList):
        out = BytesIO()
        archive = ZipFile(out, mode="w")
        manifest = Manifest()

        records = RecordList()
        xml = tostring(records.to_tree())
        records._write(archive, manifest)
        manifest.append(records)

        assert archive.namelist() == [records.path[1:]]
        assert manifest.find(records.mime_type)


    def test_extLst_round_trip(self, RecordList):
        from openpyxl.descriptors.excel import Extension, ExtensionList

        ext = ExtensionList(ext=[Extension(uri="{paper-test}")])
        records = RecordList(r=(), extLst=ext)
        node = fromstring(tostring(records.to_tree()))
        loaded = RecordList.from_tree(node)
        assert loaded.extLst is not None
        assert loaded.extLst.ext[0].uri == "{paper-test}"
        xml = tostring(loaded.to_tree())
        expected = """
        <pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0">
          <extLst>
            <ext uri="{paper-test}"/>
          </extLst>
        </pivotCacheRecords>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


def test_record_constructor_keeps_typed_fields():
    from openpyxl.pivot.fields import Boolean, DateTimeField, Error, Index, Missing, Number, Text
    from openpyxl.pivot.record import Record

    record = Record(
        m=Missing(),
        n=Number(v=2),
        b=Boolean(v=True),
        e=Error(v="#DIV/0!"),
        s=Text(v="x"),
        d=DateTimeField(v="2014-03-24T00:00:00"),
        x=Index(v=3),
    )
    assert [type(item) for item in record._fields] == [
        Missing, Number, Boolean, Error, Text, DateTimeField, Index]
    xml = tostring(record.to_tree())
    expected = """
    <r>
      <m/>
      <n v="2"/>
      <b v="1"/>
      <e v="#DIV/0!"/>
      <s v="x"/>
      <d v="2014-03-24T00:00:00"/>
      <x v="3"/>
    </r>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
    loaded = Record.from_tree(fromstring(xml))
    assert [type(item) for item in loaded._fields] == [
        Missing, Number, Boolean, Error, Text, DateTimeField, Index]
