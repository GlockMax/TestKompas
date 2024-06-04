#import "ksConstants.tlb"   no_namespace named_guids
#import "ksConstants3D.tlb" no_namespace named_guids

#import "kAPI7.tlb"

#import "kAPI5.tlb"
#include <atlbase.h>
#include <kAPI2D5COM.h>

#include <cmath>
#include <iostream>
#include <map>

struct Dimentions {
	double d0;
	double pg;
	double d;
	double p;
	int d1;
	double d2;
	double d3;
	double d4;
	double D;
	double D1;
	double L;
	double l;
	int l1;
	double b;
	double b1;
	double c;
	double c1;
	double S1;
	double scale;
	_bstr_t inches;
};

std::map<std::string, Dimentions> Table {
	std::pair<std::string, Dimentions> {"12x2", { 13.138, 1.337, 18, 1.5, 8, 13, 15.8, 11, 
	21.5, 18, 34, 14, 12, 3, 2.5, 1.6, 1.6, 19, 2, L"G 1/4"}},
	std::pair<std::string, Dimentions> {"14x2", { 16.663, 1.337, 22, 1.5, 10, 17, 19.8, 14.5, 
	27, 22, 36, 15, 12, 3, 2.5, 1.6, 1.6, 19, 2, L"G 3/8"}},
	std::pair<std::string, Dimentions> {"20x2.5", { 20.956, 1.814, 27, 1.5, 14, 22, 24.8, 18,
	34, 28, 41, 16, 14, 3, 3, 1.6, 2, 30, 2, L"G 1/2"}},
	std::pair<std::string, Dimentions> {"25x3", { 26.442, 1.814, 33, 1.5, 18, 28, 30.8, 23.9,
	41, 34, 47, 18, 16, 3, 3, 1.6, 2, 36, 1, L"G 3/4"}},
	std::pair<std::string, Dimentions> {"32x3.5", { 33.25, 2.309, 39, 1.5, 23, 34, 36.8, 29.5,
	47, 39, 52, 20, 18, 3, 4, 1.6, 2.5, 41, 1, L"G 1"}},
	std::pair<std::string, Dimentions> {"40x4", { 41.912, 2.309, 48, 1.5, 30, 42, 45.8, 38,
	56, 48, 58, 22, 20, 3, 4, 1.6, 2.5, 50, 1, L"G 1 1/4"}},
	std::pair<std::string, Dimentions> {"50x5", { 47.805, 2.309, 56, 2, 36, 48, 53, 44,
	68, 57, 64, 25, 22, 4, 4, 2, 2.5, 60, 1, L"G 1 1/2"}}
};


int main() {
	HRESULT hr;
	hr = CoInitialize(NULL);

	// Запускаем компас или заходим в открытый
	Kompas6API5::KompasObjectPtr kompas;
	hr = kompas.GetActiveObject(L"Kompas.Application.5");
	if (FAILED(hr)) kompas.CreateInstance(L"Kompas.Application.5");

	// Делаем его видимым
	kompas->Visible = true;

	// Создаём новый 3Д файл
	Kompas6API5::ksDocument3DPtr doc3DParam;
	doc3DParam = kompas->Document3D();
	doc3DParam->Create(false, true);
	doc3DParam = kompas->ActiveDocument3D(); 

	Kompas6API5::ksPartPtr part = doc3DParam->GetPart(pTop_Part);
	
	// Создаём экиз
	Kompas6API5::ksEntityPtr sketch = part->NewEntity(o3d_sketch);
	Kompas6API5::ksSketchDefinitionPtr sk_def = sketch->GetDefinition();
	// Устанавливаем для него плоскость
	sk_def->SetPlane(part->GetDefaultEntity(o3d_planeXOZ)); // YOZ
	sketch->Create();
	// Входим в режим редактирования эскиза
	Kompas6API5::ksDocument2DPtr e2d = sk_def->BeginEdit();

	// Строим эскиз

	std::string size_str = "40x4"; // 2, 2, 2, 1, 1, 1, 1
	Dimentions z = Table[size_str];

	int l1 = z.l1;
	int d1 = z.d1;
	double d0 = z.d0;
	double c1 = z.c1;
	double b1 = z.b1;
	double d4 = z.d4;
	double D = z.D;
	double D1 = z.D1;
	double l = z.l;
	double L = z.L;
	double d3 = z.d3;
	double b = z.b;
	double d = z.d;
	double c = z.c;
	double d2 = z.d2;
	double degrees = 37;
	double S1 = z.S1;
	double p = z.p;
	double pg = z.pg;
	_bstr_t inches = z.inches;
	double scale = z.scale;


	// Осевая
	e2d->ksLineSeg(-l1, 0, L-l1, 0, 3);

	e2d->ksLineSeg(-l1, d1 / 2, -l1, (d0 / 2) - c1, 1);
	e2d->ksLineSeg(-l1, (d0 / 2) - c1, -l1 + c1, d0 / 2, 1);

	double c2 = d0 / 2 - d4 / 2;
	e2d->ksLineSeg(-l1 + c1, d0 / 2, -b1 - c2, d0 / 2, 1);
	e2d->ksLineSeg(-b1 - c2, d0 / 2, -b1, d4 / 2, 1);
	e2d->ksLineSeg(-b1, d4 / 2, 0, d4/2, 1);
	e2d->ksLineSeg(0, d4 / 2, 0, D/2, 1);

	double c3 = D / 2 - D1 / 2;
	double l2 = L - l - l1;
	e2d->ksLineSeg(0, D / 2, l2 - c3 / 1.732, D / 2, 1);
	e2d->ksLineSeg(l2 - c3 / 1.732, D / 2, l2, D1/2, 1);

	e2d->ksLineSeg(l2, D1 / 2, l2, d3 / 2, 1);
	e2d->ksLineSeg(l2, d3 / 2, l2 + b, d3 / 2, 1);

	double c4 = d / 2 - d3/2;
	e2d->ksLineSeg(l2 + b, d3 / 2, l2 + b + c4, d / 2, 1);
	e2d->ksLineSeg(l2 + b + c4, d / 2, l2 + l - c, d / 2, 1);
	e2d->ksLineSeg(l2 + l - c, d / 2, l2 + l, d / 2-c, 1);
	e2d->ksLineSeg(l2 + l, d / 2 - c, l2 + l, d2 / 2, 1);

	e2d->ksLineSeg(l2 + l, d2 / 2, l2 + l - ((d2 - d1) / 2) / tan(degrees * (3.1415926 / 180) / 2), d1 / 2, 1);
	e2d->ksLineSeg(l2 + l - ((d2 - d1) / 2) / tan(degrees * (3.1415926 / 180) / 2), d1 / 2, -l1, d1 / 2, 1);


	// Выходим из режима эскиза
	sk_def->EndEdit();

	// Операция вращения
	Kompas6API5::ksEntityPtr bossRot = part->NewEntity(o3d_bossRotated);
	Kompas6API5::ksBossRotatedDefinitionPtr br_def = bossRot->GetDefinition();
	br_def->SetSideParam(TRUE, 360);
	br_def->directionType =  (dtNormal);
	br_def->SetSketch(sketch);
	bossRot->Create();

	Kompas6API5::ksEntityPtr sketch1 = part->NewEntity(o3d_sketch);
	Kompas6API5::ksSketchDefinitionPtr sk_def1 = sketch1->GetDefinition();
	// Устанавливаем для него плоскость
	sk_def1->SetPlane(part->GetDefaultEntity(o3d_planeYOZ)); // XOY
	sketch1->Create();
	// Входим в режим редактирования эскиза
	Kompas6API5::ksDocument2DPtr e2d1 = sk_def1->BeginEdit();

	// Задаём параметры шестиугольника
	Kompas6API5::ksRegularPolygonParamPtr par = kompas->GetParamStruct(ko_RegularPolygonParam);
	par->count = 6;
	par->xc = 0; par->yc = 0;
	par->style = 1;
	par->ang = 30; par->describe = TRUE;
	par->radius = S1 / 2;
	// Рисуем его
	e2d1->ksRegularPolygon(par, 0);

	// Окружность для корректной операции вырезания
	e2d1->ksCircle(0, 0, D + 2, 1);

	// Выходим из режима эскиза
	sk_def1->EndEdit();

	// Вырзание выдавливанием
	Kompas6API5::ksEntityPtr cut = part->NewEntity(o3d_cutExtrusion);
	Kompas6API5::ksCutExtrusionDefinitionPtr cut_def = cut->GetDefinition();
	cut_def->SetSideParam(TRUE, etBlind, l2, 0, FALSE);
	cut_def->SetSketch(sketch1);
	cut_def->directionType = (dtNormal);
	cut->Create();

	// Самое весёлое - обозначение резьбы
	// Метрическая:
	Kompas6API5::ksEntityPtr threadm = part->NewEntity(o3d_thread);
	Kompas6API5::ksThreadDefinitionPtr th_def = threadm->GetDefinition();
	th_def->allLength = TRUE;
	th_def->dr = d; th_def->autoDefinDr = FALSE;
	th_def->p = p;

	// Дюймовая:
	Kompas6API5::ksEntityPtr threadg = part->NewEntity(o3d_thread);
	Kompas6API5::ksThreadDefinitionPtr thg_def = threadg->GetDefinition();
	thg_def->allLength = TRUE;
	thg_def->dr = d0; thg_def->autoDefinDr = FALSE;
	thg_def->p = pg;

	// Ищем грань, куда положить резьбы

	// Наше тело:
	Kompas6API5::ksBodyPtr bp = part->GetMainBody();
	Kompas6API5::ksFaceCollectionPtr fc = bp->FaceCollection();
	
	for (int i = fc->GetCount()-1; i > 0; i--) {
		Kompas6API5::ksFaceDefinitionPtr t = fc->GetByIndex(i);
		if (t->IsCylinder() == -1) {
			double r, h;
			t->GetCylinderParam(&h, &r);
			if (r == d / 2) {
				int q = th_def->SetBaseObject(t->GetEntity());
				std::cout << "Setting face metric: " << q << std::endl;
			}
			else if (r == d0 / 2) {
				int q = thg_def->SetBaseObject(t->GetEntity());
				std::cout << "Setting face imperic: " << q << std::endl;
			}
		}
		t->Release();
	}

	threadm->Create();
	threadg->Create();

	// Сохраняем в файл
	_bstr_t file = L"C:\\Users\\Администратор\\Desktop\\Штуцер.m3d";
	doc3DParam->SaveAs(file);

	// ===============================================
	//       ||||||||||||||        |||||||||||||||
	// ================================================

	// Получаем 7 API
	KompasAPI7::IApplicationPtr kompas7 = kompas->ksGetApplication7();

	// Коллекция документов, открытых в приложении
	KompasAPI7::IDocumentsPtr dd = kompas7->Documents;

	// Создаём дефолтный чертёж по сохранённым настройкам пользователя
	KompasAPI7::IKompasDocument2DPtr d2dd = dd->AddWithDefaultSettings(ksDocumentDrawing, true);

	// Неуказанная шероховатость
	KompasAPI7::ISpecRoughPtr sr = ((KompasAPI7::IDrawingDocumentPtr)d2dd)->SpecRough;
	sr->AddSign = TRUE;
	sr->AutoPlacement = TRUE;
	sr->SignType = ksRoughSignEnum::ksNoProcessingType;
	sr->Text = (_bstr_t)L"Ra6,3";
	sr->Update();

	// Достаём менеджера видов и слоёв
	KompasAPI7::IViewsAndLayersManagerPtr pViewM = d2dd->GetViewsAndLayersManager();
	KompasAPI7::IViewsPtr views = pViewM->Views; // массив видов


	const _variant_t projTypes[2] // = {ksVPFront, ksVPLeft, ksVPUp};
	{ {(long)5, (VARTYPE)VT_I4}, {(long)1, (VARTYPE)VT_I4} };

	double X0 = 70; double Y0 = 190;
	double DX = 40; double DY = DX;

	views->AddStandartViews(file,
		L"#Слева", projTypes, X0, Y0, scale, DX, DY);

	// Наши виды
	KompasAPI7::IViewPtr view_main = views->GetActiveView(); // Активный вид у нас главный вид
	KompasAPI7::IViewPtr view_sub = views->View[2];

	KompasAPI7::IDrawingContainerPtr draw_cont = view_main;
	KompasAPI7::IRectanglesPtr rects = draw_cont->Rectangles;
	KompasAPI7::IRectanglePtr rect = rects->Add();
	rect->Style = 1;
	rect->Height = D;
	rect->Width = L;
	rect->X = -l1;
	rect->Y = -D/2;
	rect->Update();

	KompasAPI7::ICutViewParamPtr cv = view_main;
	std::cout << cv->AddCut((_bstr_t)L"Разрезик", 12, X0 + l2 + l + DX + S1/2, Y0, 
		FALSE, rect, view_sub) << std::endl;
	// X и Y В СИСТЕМНЫХ КООРДИНАТАХ

	KompasAPI7::ISymbols2DContainerPtr sizes_main = view_main;
	KompasAPI7::IAngleDimensionsPtr ad = sizes_main->AngleDimensions; // Угловые размеры
	KompasAPI7::ILineDimensionsPtr ld = sizes_main->LineDimensions; // Линейные размеры

	KompasAPI7::IAngleDimensionPtr adim = ad->Add(ksDrADimension);
	adim->Angle1 = 360 - ((double)37) / 2;
	adim->Angle2 = ((double)37) / 2;
	adim->Radius = -((-(d2 / 2) / tan(37 * (3.1415926 / 180) / 2)) + l2 + l) + l2 + l - (((d2 - d1) / 2) / cos(degrees * (3.1415926 / 180) / 2)) / 2;
	adim->Direction = 1; adim->DimensionType = ksAngleDimTypeEnum::ksADMinAngle;
	adim->X1 = l2 + l - ((d2 - d1) / 2) / tan(degrees * (3.1415926 / 180) / 2); adim->Y1 = -((double) d1)/2;
	adim->X2 = l2 + l - ((d2 - d1) / 2) / tan(degrees * (3.1415926 / 180) / 2); adim->Y2 = ((double) d1)/2;
	adim->Xc = (-(d2/2) / tan(37 * (3.1415926 / 180) / 2)) + l2 + l; adim->Yc = 0;
	adim->X3 = l2 + l - (((d2 - d1) / 2) / cos(degrees * (3.1415926 / 180) / 2))/2; adim->Y3 = 0;
	// центр окружности расситаем по формуле d2 / tan(degrees * (3.1415926 / 180))
	adim->Update();

	adim = ad->Add(ksDrADimension);
	adim->Angle1 = 90;
	adim->Angle2 = 135;
	adim->Radius = 15;
	adim->Direction = 1; adim->DimensionType = ksAngleDimTypeEnum::ksADMinAngle;
	adim->X1 = -b1; adim->Y1 = d4/2;
	adim->X2 = -b1; adim->Y2 = d4/2;
	adim->Xc = -b1; adim->Yc = d4/2;
	adim->X3 = -b1 - 15 * sin(45 * (3.1415926 / 180) / 2); adim->Y3 = d4/2 + 15 * cos(45 * (3.1415926 / 180) / 2);

	KompasAPI7::IDimensionParamsPtr adim_p = adim;
	adim_p->ArrowPos = ksDimensionArrowPosEnum::ksDimArrowOutside;

	adim->Update();

	adim = ad->Add(ksDrADimension);
	adim->Angle1 = 45;
	adim->Angle2 = 90;
	adim->Radius = 15;
	adim->Direction = 1; adim->DimensionType = ksAngleDimTypeEnum::ksADMinAngle;
	adim->X1 = l2+b; adim->Y1 = d3/2;
	adim->X2 = l2 + b; adim->Y2 = d3 / 2;
	adim->Xc = l2 + b; adim->Yc = d3 / 2;
	adim->X3 = l2 + b + 15 * sin(45 * (3.1415926 / 180) / 2); adim->Y3 = d3 / 2 + 15 * cos(45 * (3.1415926 / 180) / 2);

	adim_p = adim;
	adim_p->ArrowPos = ksDimensionArrowPosEnum::ksDimArrowOutside;

	adim->Update();

	// L
	KompasAPI7::ILineDimensionPtr linedim = ld->Add();
	linedim->X1 = -l1; linedim->X2 = l2 + l; linedim->X3 = 0;
	linedim->Y1 = 0; linedim->Y2 = 0; linedim->Y3 = -D/2 - 50;
	linedim->Update();

	// d1
	linedim = ld->Add();
	linedim->X1 = -l1; linedim->X2 = -l1; linedim->X3 = -l1-10;
	linedim->Y1 = -((double)d1)/2; linedim->Y2 = ((double)d1) / 2; linedim->Y3 = 0;
	KompasAPI7::IDimensionTextPtr linedim_t = linedim;
	linedim_t->Sign = 1;
	linedim->Update();

	// d
	linedim = ld->Add();
	linedim->X1 = l2 + l - c; linedim->X2 = l2 + l - c; linedim->X3 = l2+l + 20;
	linedim->Y1 = d / 2; linedim->Y2 = -d / 2; linedim->Y3 = 0;

	linedim_t = linedim;
	linedim_t->Sign = 4;
	KompasAPI7::ITextLinePtr suff = linedim_t->Suffix;
	KompasAPI7::ITextItemPtr sf = suff->Add();
	sf->Str = (_bstr_t)L"x1,5";
	sf->Update();
	linedim->Update();

	// d2
	linedim = ld->Add();
	linedim->X1 = l2 + l; linedim->X2 = l2 + l; linedim->X3 = l2+l + 10;
	linedim->Y1 = -((double)d2) / 2; linedim->Y2 = ((double)d2) / 2; linedim->Y3 = 0;
	linedim_t = linedim;
	linedim_t->Sign = 1;
	linedim->Update();

	// c1
	linedim = ld->Add();
	linedim->X1 = -l1; linedim->X2 = -l1 + c1; linedim->X3 = -l1 - c1;
	linedim->Y1 = -d0 / 2; linedim->Y2 = -d0 / 2; linedim->Y3 = -D/2 - 10;

	KompasAPI7::IDimensionParamsPtr linedim_p = linedim;
	linedim_p->ArrowPos = ksDimensionArrowPosEnum::ksDimArrowOutside;

	linedim_t = linedim;
	suff = linedim_t->Suffix;
	sf = suff->Add();
	sf->Str = (_bstr_t)L"x45°";
	sf->Update();
	linedim->Update();

	// c
	linedim = ld->Add();
	linedim->X1 = l2 + l; linedim->X2 = l2 + l - c; linedim->X3 = l2+l+c;
	linedim->Y1 = -d / 2; linedim->Y2 = -d / 2; linedim->Y3 = -D / 2 - 10;

	linedim_p = linedim;
	linedim_p->ArrowPos = ksDimensionArrowPosEnum::ksDimArrowOutside;

	linedim_t = linedim;
	suff = linedim_t->Suffix;
	sf = suff->Add();
	sf->Str = (_bstr_t)L"x45°";
	sf->Update();
	linedim->Update();

	// l1
	linedim = ld->Add();
	linedim->X1 = -l1; linedim->X2 = 0; linedim->X3 = -l1/2;
	linedim->Y1 = -((double)d4) / 2; linedim->Y2 = -((double)d4) / 2; linedim->Y3 = -D/2-40;
	linedim->Update();
	
	// l
	linedim = ld->Add();
	linedim->X1 = l2; linedim->X2 = l2+l; linedim->X3 = l2+l / 2;
	linedim->Y1 = -((double)d3) / 2; linedim->Y2 = -((double)d3) / 2; linedim->Y3 = -D / 2 - 40;
	linedim->Update();

	// b1
	linedim = ld->Add();
	linedim->X1 = -b1; linedim->X2 = 0; linedim->X3 = -b1-b1;
	linedim->Y1 = -((double)d4) / 2; linedim->Y2 = -((double)d4) / 2; linedim->Y3 = -D / 2 - 20;

	linedim_p = linedim;
	linedim_p->ArrowPos = ksDimensionArrowPosEnum::ksDimArrowOutside;
	linedim->Update();

	// b
	linedim = ld->Add();
	linedim->X1 = l2; linedim->X2 = l2+b; linedim->X3 = l2+b+b;
	linedim->Y1 = -((double)d3) / 2; linedim->Y2 = -((double)d3) / 2; linedim->Y3 = -D / 2 - 30;

	linedim_p = linedim;
	linedim_p->ArrowPos = ksDimensionArrowPosEnum::ksDimArrowOutside;
	linedim->Update();

	// d4
	linedim = ld->Add();
	linedim->X1 = -2*b1/3; linedim->X2 = -2 * b1 / 3; linedim->X3 = -2 * b1 / 3;
	linedim->Y1 = -d4 / 2; linedim->Y2 = d4 / 2; linedim->Y3 = 0;

	linedim_t = linedim;
	linedim_t->Sign = 1;
	linedim->Update();

	// D1
	linedim = ld->Add();
	linedim->X1 = l2; linedim->X2 = l2; linedim->X3 = 4*l2/8;
	linedim->Y1 = -D1 / 2; linedim->Y2 = D1 / 2; linedim->Y3 = 0;

	linedim_t = linedim;
	linedim_t->Sign = 1;
	linedim->Update();

	// d3
	linedim = ld->Add();
	linedim->X1 = l2+b; linedim->X2 = l2+b; linedim->X3 = l2+b;
	linedim->Y1 = -d3 / 2; linedim->Y2 = d3 / 2; linedim->Y3 = 0;

	linedim_t = linedim;
	linedim_t->Sign = 1;
	linedim->Update();

	// Дюймовая резьба
	KompasAPI7::ILeadersPtr lead = sizes_main->Leaders;
	KompasAPI7::IBaseLeaderPtr lgb = lead->Add(ksDrLeader);
	KompasAPI7::ILeaderPtr lg = lgb;
	lg->ShelfDirection = ksShelfDirectionEnum::ksLSLeft;

	KompasAPI7::IBranchsPtr brs = lgb;
	brs->X0 = -l1-10; brs->Y0 = D / 2;
	brs->AddBranchByPoint(2, -l1+(3*c1/2), d0 / 2);

	KompasAPI7::ITextPtr text = lg->TextOnShelf;
	KompasAPI7::ITextLinePtr text1 = text->Add();
	KompasAPI7::ITextItemPtr text2 = text1->Add();
	text2->Str = inches;
	text2->Update();

	lgb->Update();

	// На второй вид
	KompasAPI7::ISymbols2DContainerPtr sizes_sub = view_sub;
	KompasAPI7::ILineDimensionsPtr ldb = sizes_sub->LineDimensions; // Линейные размеры

	// D
	linedim = ldb->Add();
	linedim->X1 = 0; linedim->X2 = 0; linedim->X3 = D/2 + 10;
	linedim->Y1 = -D / 2; linedim->Y2 = D / 2; linedim->Y3 = 0;

	linedim_t = linedim;
	linedim_t->Sign = 1;
	linedim->Update();

	// S1
	linedim = ldb->Add();
	linedim->X1 = -S1/2; linedim->X2 = S1/2; linedim->X3 = 0;
	linedim->Y1 = 0; linedim->Y2 = 0; linedim->Y3 = D / 2 + 10;
	linedim->Update();


	kompas->Release();
	CoUninitialize();
	return 0;
}