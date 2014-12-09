package oleutil

import (
	"github.com/mattn/go-ole"
	"math"
	"testing"
)

const EPSILON = 1.19e-7

func newTestObject(t *testing.T) *ole.IDispatch {
	ole.CoInitialize(0)

	unk, err := CreateObject("GoOleTests.Tests")
	if err != nil {
		t.Log(err)
		t.FailNow()
	}
	disp, _ := unk.QueryInterface(ole.IID_IDispatch)
	unk.Release()
	return disp
}

func TestCurrency(t *testing.T) {
	testobj := newTestObject(t)
	v := MustCallMethod(testobj, "TestReturnCurrency", ole.Currency(2134.5967)).Value().(ole.Currency)
	if !float64Equal(float64(v), 2134.5967) {
		t.Log("TestCurrency returned", v, "but expected 2134.5967")
		t.FailNow()
	}
}

func TestCurrencyByRef(t *testing.T) {
	testobj := newTestObject(t)
	c := ole.Currency(2134.5967)
	v := MustCallMethod(testobj, "TestReturnCurrencyByRef", &c).Value().(ole.Currency)
	if !float64Equal(float64(v), 2134.5967) {
		t.Log("TestCurrency returned", v, "but expected 2134.5967")
		t.FailNow()
	}
}

func TestFloat32(t *testing.T) {
	testobj := newTestObject(t)
	v := MustCallMethod(testobj, "TestReturnFloat32", 5.4321).Value().(float32)
	if !float32Equal(v, 5.4321) {
		t.Log("TestReturnsFloat32 returned", v, "but expected 5.4321")
		t.FailNow()
	}
}

func TestFloat64(t *testing.T) {
	testobj := newTestObject(t)
	v := MustCallMethod(testobj, "TestReturnFloat64", 5.4321).Value().(float64)
	if !float64Equal(v, 5.4321) {
		t.Log("TestReturnsFloat64 returned", v, "but expected 5.4321")
		t.FailNow()
	}
}

func TestFloat32ByRef(t *testing.T) {
	testobj := newTestObject(t)
	f := float32(5.4321)
	v := MustCallMethod(testobj, "TestReturnFloat32ByRef", &f).Value().(float32)
	if !float32Equal(v, 5.4321) {
		t.Log("TestReturnsFloat32ByRef returned", v, "but expected 5.4321")
		t.FailNow()
	}
}

func TestFloat64ByRef(t *testing.T) {
	testobj := newTestObject(t)
	f := float64(5.4321)
	v := MustCallMethod(testobj, "TestReturnFloat64ByRef", &f).Value().(float64)
	if !float64Equal(v, 5.4321) {
		t.Log("TestFloat64ByRef returned", v, "but expected 5.4321")
		t.FailNow()
	}
}

func TestFloat32String(t *testing.T) {
	testobj := newTestObject(t)
	v := MustCallMethod(testobj, "TestReturnFloat32", "5.4321").Value().(float32)
	if !float32Equal(v, 5.4321) {
		t.Log("TestReturnsFloat32 returned", v, "but expected 5.4321")
		t.FailNow()
	}
}

func TestFloat64String(t *testing.T) {
	testobj := newTestObject(t)
	v := MustCallMethod(testobj, "TestReturnFloat64", "5.4321").Value().(float64)
	if !float64Equal(v, 5.4321) {
		t.Log("TestReturnsFloat64 returned", v, "but expected 5.4321")
		t.FailNow()
	}
}

func float32Equal(f1 float32, f2 float32) bool {
	v := math.Abs(float64(f1) - float64(f2))
	return v <= EPSILON
}

func float64Equal(f1 float64, f2 float64) bool {
	v := math.Abs(f1 - f2)
	return v <= EPSILON
}
