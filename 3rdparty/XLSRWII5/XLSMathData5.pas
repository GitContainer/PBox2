unit XLSMathData5;

{$I AxCompilers.inc}

interface

uses SysUtils, Math,
     XLSUtils5;

const MAX_QUICKFACTORIAL = 170;
  const_e = 2.7182818284590452353602874713526624977572470936999595749669676;
  const_epsilon = 2.4651903288156618919116517665087e-32;
  MaxSeriesIterations = 100000;
  ResNumDigits = 53;
  LogMaxValue = 709;
  LogMinValue = -707;
  NiceArraySize = 30;
  IsSpecialized = True;

function QuickFactorial(AValue: integer): double;
function ErfImp(z: double; invert: boolean): double;
function erfc(z: double): double;
function expm1(x: double): double;
function log1p(x: double): double; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
function powm1(a,z: double): double; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
function itrunc(V: double): integer; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

function EvaluatePolynomial3(const a: array of double; const x: double): double;
function EvaluatePolynomial4(const a: array of double; const x: double): double;
function EvaluatePolynomial5(const a: array of double; const x: double): double;
function EvaluatePolynomial6(const a: array of double; const x: double): double;
function EvaluatePolynomial7(const a: array of double; const x: double): double;
function EvaluatePolynomial8(const a: array of double; const x: double): double;
function EvaluatePolynomial9(const a: array of double; const x: double): double;
function EvaluatePolynomial10(const a: array of double; const x: double): double;
function EvaluatePolynomial11(const a: array of double; const x: double): double;

function EvaluateEvenPolynomial3(const a: array of double; const x: double): double;
function EvaluateEvenPolynomial4(const a: array of double; const x: double): double;
function EvaluateEvenPolynomial5(const a: array of double; const x: double): double;
function EvaluateEvenPolynomial6(const a: array of double; const x: double): double;
function EvaluateEvenPolynomial7(const a: array of double; const x: double): double;
function EvaluateEvenPolynomial8(const a: array of double; const x: double): double;

function EvaluateOddPolynomial11(poly: array of double; z: double): double;

implementation

const
  QFactorials: array[0..170] of double = (
      1,
      1,
      2,
      6,
      24,
      120,
      720,
      5040,
      40320,
      362880.0,
      3628800.0,
      39916800.0,
      479001600.0,
      6227020800.0,
      87178291200.0,
      1307674368000.0,
      20922789888000.0,
      355687428096000.0,
      6402373705728000.0,
      121645100408832000.0,
      0.243290200817664e19,
      0.5109094217170944e20,
      0.112400072777760768e22,
      0.2585201673888497664e23,
      0.62044840173323943936e24,
      0.15511210043330985984e26,
      0.403291461126605635584e27,
      0.10888869450418352160768e29,
      0.304888344611713860501504e30,
      0.8841761993739701954543616e31,
      0.26525285981219105863630848e33,
      0.822283865417792281772556288e34,
      0.26313083693369353016721801216e36,
      0.868331761881188649551819440128e37,
      0.29523279903960414084761860964352e39,
      0.103331479663861449296666513375232e41,
      0.3719933267899012174679994481508352e42,
      0.137637530912263450463159795815809024e44,
      0.5230226174666011117600072241000742912e45,
      0.203978820811974433586402817399028973568e47,
      0.815915283247897734345611269596115894272e48,
      0.3345252661316380710817006205344075166515e50,
      0.1405006117752879898543142606244511569936e52,
      0.6041526306337383563735513206851399750726e53,
      0.265827157478844876804362581101461589032e55,
      0.1196222208654801945619631614956577150644e57,
      0.5502622159812088949850305428800254892962e58,
      0.2586232415111681806429643551536119799692e60,
      0.1241391559253607267086228904737337503852e62,
      0.6082818640342675608722521633212953768876e63,
      0.3041409320171337804361260816606476884438e65,
      0.1551118753287382280224243016469303211063e67,
      0.8065817517094387857166063685640376697529e68,
      0.427488328406002556429801375338939964969e70,
      0.2308436973392413804720927426830275810833e72,
      0.1269640335365827592596510084756651695958e74,
      0.7109985878048634518540456474637249497365e75,
      0.4052691950487721675568060190543232213498e77,
      0.2350561331282878571829474910515074683829e79,
      0.1386831185456898357379390197203894063459e81,
      0.8320987112741390144276341183223364380754e82,
      0.507580213877224798800856812176625227226e84,
      0.3146997326038793752565312235495076408801e86,
      0.1982608315404440064116146708361898137545e88,
      0.1268869321858841641034333893351614808029e90,
      0.8247650592082470666723170306785496252186e91,
      0.5443449390774430640037292402478427526443e93,
      0.3647111091818868528824985909660546442717e95,
      0.2480035542436830599600990418569171581047e97,
      0.1711224524281413113724683388812728390923e99,
      0.1197857166996989179607278372168909873646e101,
      0.8504785885678623175211676442399260102886e102,
      0.6123445837688608686152407038527467274078e104,
      0.4470115461512684340891257138125051110077e106,
      0.3307885441519386412259530282212537821457e108,
      0.2480914081139539809194647711659403366093e110,
      0.188549470166605025498793226086114655823e112,
      0.1451830920282858696340707840863082849837e114,
      0.1132428117820629783145752115873204622873e116,
      0.8946182130782975286851441715398316520698e117,
      0.7156945704626380229481153372318653216558e119,
      0.5797126020747367985879734231578109105412e121,
      0.4753643337012841748421382069894049466438e123,
      0.3945523969720658651189747118012061057144e125,
      0.3314240134565353266999387579130131288001e127,
      0.2817104114380550276949479442260611594801e129,
      0.2422709538367273238176552320344125971528e131,
      0.210775729837952771721360051869938959523e133,
      0.1854826422573984391147968456455462843802e135,
      0.1650795516090846108121691926245361930984e137,
      0.1485715964481761497309522733620825737886e139,
      0.1352001527678402962551665687594951421476e141,
      0.1243841405464130725547532432587355307758e143,
      0.1156772507081641574759205162306240436215e145,
      0.1087366156656743080273652852567866010042e147,
      0.103299784882390592625997020993947270954e149,
      0.9916779348709496892095714015418938011582e150,
      0.9619275968248211985332842594956369871234e152,
      0.942689044888324774562618574305724247381e154,
      0.9332621544394415268169923885626670049072e156,
      0.9332621544394415268169923885626670049072e158,
      0.9425947759838359420851623124482936749562e160,
      0.9614466715035126609268655586972595484554e162,
      0.990290071648618040754671525458177334909e164,
      0.1029901674514562762384858386476504428305e167,
      0.1081396758240290900504101305800329649721e169,
      0.1146280563734708354534347384148349428704e171,
      0.1226520203196137939351751701038733888713e173,
      0.132464181945182897449989183712183259981e175,
      0.1443859583202493582204882102462797533793e177,
      0.1588245541522742940425370312709077287172e179,
      0.1762952551090244663872161047107075788761e181,
      0.1974506857221074023536820372759924883413e183,
      0.2231192748659813646596607021218715118256e185,
      0.2543559733472187557120132004189335234812e187,
      0.2925093693493015690688151804817735520034e189,
      0.339310868445189820119825609358857320324e191,
      0.396993716080872089540195962949863064779e193,
      0.4684525849754290656574312362808384164393e195,
      0.5574585761207605881323431711741977155627e197,
      0.6689502913449127057588118054090372586753e199,
      0.8094298525273443739681622845449350829971e201,
      0.9875044200833601362411579871448208012564e203,
      0.1214630436702532967576624324188129585545e206,
      0.1506141741511140879795014161993280686076e208,
      0.1882677176888926099743767702491600857595e210,
      0.237217324288004688567714730513941708057e212,
      0.3012660018457659544809977077527059692324e214,
      0.3856204823625804217356770659234636406175e216,
      0.4974504222477287440390234150412680963966e218,
      0.6466855489220473672507304395536485253155e220,
      0.8471580690878820510984568758152795681634e222,
      0.1118248651196004307449963076076169029976e225,
      0.1487270706090685728908450891181304809868e227,
      0.1992942746161518876737324194182948445223e229,
      0.269047270731805048359538766214698040105e231,
      0.3659042881952548657689727220519893345429e233,
      0.5012888748274991661034926292112253883237e235,
      0.6917786472619488492228198283114910358867e237,
      0.9615723196941089004197195613529725398826e239,
      0.1346201247571752460587607385894161555836e242,
      0.1898143759076170969428526414110767793728e244,
      0.2695364137888162776588507508037290267094e246,
      0.3854370717180072770521565736493325081944e248,
      0.5550293832739304789551054660550388118e250,
      0.80479260574719919448490292577980627711e252,
      0.1174997204390910823947958271638517164581e255,
      0.1727245890454638911203498659308620231933e257,
      0.2556323917872865588581178015776757943262e259,
      0.380892263763056972698595524350736933546e261,
      0.571338395644585459047893286526105400319e263,
      0.8627209774233240431623188626544191544816e265,
      0.1311335885683452545606724671234717114812e268,
      0.2006343905095682394778288746989117185662e270,
      0.308976961384735088795856467036324046592e272,
      0.4789142901463393876335775239063022722176e274,
      0.7471062926282894447083809372938315446595e276,
      0.1172956879426414428192158071551315525115e279,
      0.1853271869493734796543609753051078529682e281,
      0.2946702272495038326504339507351214862195e283,
      0.4714723635992061322406943211761943779512e285,
      0.7590705053947218729075178570936729485014e287,
      0.1229694218739449434110178928491750176572e290,
      0.2004401576545302577599591653441552787813e292,
      0.3287218585534296227263330311644146572013e294,
      0.5423910666131588774984495014212841843822e296,
      0.9003691705778437366474261723593317460744e298,
      0.1503616514864999040201201707840084015944e301,
      0.2526075744973198387538018869171341146786e303,
      0.4269068009004705274939251888899566538069e305,
      0.7257415615307998967396728211129263114717e307);

function QuickFactorial(AValue: integer): double;
begin
  Result := QFactorials[AValue];
end;

function EvaluatePolynomial3(const a: array of double; const x: double): double;
begin
   Result := (a[2] * x + a[1]) * x + a[0];
end;

function EvaluatePolynomial4(const a: array of double; const x: double): double;
begin
   Result := ((a[3] * x + a[2]) * x + a[1]) * x + a[0];
end;

function EvaluatePolynomial5(const a: array of double; const x: double): double;
begin
   Result := (((a[4] * x + a[3]) * x + a[2]) * x + a[1]) * x + a[0];
end;

function EvaluatePolynomial6(const a: array of double; const x: double): double;
begin
   Result := ((((a[5] * x + a[4]) * x + a[3]) * x + a[2]) * x + a[1]) * x + a[0];
end;

function EvaluatePolynomial7(const a: array of double; const x: double): double;
begin
   Result := (((((a[6] * x + a[5]) * x + a[4]) * x + a[3]) * x + a[2]) * x + a[1]) * x + a[0];
end;

function EvaluatePolynomial8(const a: array of double; const x: double): double;
begin
   Result := ((((((a[7] * x + a[6]) * x + a[5]) * x + a[4]) * x + a[3]) * x + a[2]) * x + a[1]) * x + a[0];
end;

function EvaluatePolynomial9(const a: array of double; const x: double): double;
begin
   Result := (((((((a[8] * x + a[7]) * x + a[6]) * x + a[5]) * x + a[4]) * x + a[3]) * x + a[2]) * x + a[1]) * x + a[0];
end;

function EvaluatePolynomial10(const a: array of double; const x: double): double;
begin
   Result := ((((((((a[9] * x + a[8]) * x + a[7]) * x + a[6]) * x + a[5]) * x + a[4]) * x + a[3]) * x + a[2]) * x + a[1]) * x + a[0];
end;

function EvaluatePolynomial11(const a: array of double; const x: double): double;
begin
   Result := (((((((((a[10] * x + a[9]) * x + a[8]) * x + a[7]) * x + a[6]) * x + a[5]) * x + a[4]) * x + a[3]) * x + a[2]) * x + a[1]) * x + a[0];
end;

function EvaluateEvenPolynomial3(const a: array of double; const x: double): double;
begin
  Result := EvaluatePolynomial3(a,x * x);
end;

function EvaluateEvenPolynomial4(const a: array of double; const x: double): double;
begin
  Result := EvaluatePolynomial4(a,x * x);
end;

function EvaluateEvenPolynomial5(const a: array of double; const x: double): double;
begin
  Result := EvaluatePolynomial5(a,x * x);
end;

function EvaluateEvenPolynomial6(const a: array of double; const x: double): double;
begin
  Result := EvaluatePolynomial6(a,x * x);
end;

function EvaluateEvenPolynomial7(const a: array of double; const x: double): double;
begin
  Result := EvaluatePolynomial7(a,x * x);
end;

function EvaluateEvenPolynomial8(const a: array of double; const x: double): double;

begin

  Result := EvaluatePolynomial8(a,x * x);
end;


function EvaluateOddPolynomial11(poly: array of double; z: double): double;
begin
   Result := poly[0] + z * EvaluatePolynomial10(poly[1],z * z);
end;

function ErfImp(z: double; invert: boolean): double;
var
  Y: double;
  P: array[0..6] of double;
  Q: array[0..6] of double;
begin
   if z < 0 then begin
      if (not invert) then
         Result := -ErfImp(-z, invert)
      else if (z < -0.5) then
         Result := 2 - ErfImp(-z, invert)
      else
         Result := 1 + ErfImp(-z, false);
      Exit;
   end;
   //
   // Big bunch of selection statements now to pick
   // which implementation to use,
   // try to put most likely options first:
   //
   if (z < 0.5) then begin
      //
      // We're going to calculate erf:
      //
      if (z = 0) then
         Result := 0
      else if (z < 1e-10) then
         Result := (z * 1.125 + z * 0.003379167095512573896158903121545171688)
      else begin
         // Maximum Deviation Found:                     1.561e-17
         // Expected Error Term:                         1.561e-17
         // Maximum Relative Change in Control Points:   1.155e-04
         // Max Error found at double precision =        2.961182e-17

         Y := 1.044948577880859375;

         P[0] := 0.0834305892146531832907;
         P[1] := -0.338165134459360935041;
         P[2] := -0.0509990735146777432841;
         P[3] := -0.00772758345802133288487;
         P[4] := -0.000322780120964605683831;

         Q[0] := 1;
         Q[1] := 0.455004033050794024546;
         Q[2] := 0.0875222600142252549554;
         Q[3] := 0.00858571925074406212772;
         Q[4] := 0.000370900071787748000569;

         Result := z * (Y + EvaluatePolynomial5(P, z * z) / EvaluatePolynomial5(Q, z * z));
      end
   end
   else if ((z < 14) or ((z < 28) and invert)) then begin
      //
      // We'll be calculating erfc:
      //
      invert := not invert;
      if (z < 1.5) then begin
         // Maximum Deviation Found:                     3.702e-17
         // Expected Error Term:                         3.702e-17
         // Maximum Relative Change in Control Points:   2.845e-04
         // Max Error found at double precision =        4.841816e-17
         Y := 0.405935764312744140625;

         P[0] := -0.098090592216281240205;
         P[1] := 0.178114665841120341155;
         P[2] := 0.191003695796775433986;
         P[3] := 0.0888900368967884466578;
         P[4] := 0.0195049001251218801359;
         P[5] := 0.00180424538297014223957;

         Q[0] := 1;
         Q[1] := 1.84759070983002217845;
         Q[2] := 1.42628004845511324508;
         Q[3] := 0.578052804889902404909;
         Q[4] := 0.12385097467900864233;
         Q[5] := 0.0113385233577001411017;
         Q[6] := 0.337511472483094676155e-5;

         Result := Y + EvaluatePolynomial7(P, z - 0.5) / EvaluatePolynomial7(Q, z - 0.5);
         Result := Result * (exp(-z * z) / z);
      end
      else if (z < 2.5) then begin
         // Max Error found at double precision =        6.599585e-18
         // Maximum Deviation Found:                     3.909e-18
         // Expected Error Term:                         3.909e-18
         // Maximum Relative Change in Control Points:   9.886e-05
         Y := 0.50672817230224609375;

         P[0] := -0.0243500476207698441272;
         P[1] := 0.0386540375035707201728;
         P[2] := 0.04394818964209516296;
         P[3] := 0.0175679436311802092299;
         P[4] := 0.00323962406290842133584;
         P[5] := 0.000235839115596880717416;

         Q[0] := 1;
         Q[1] := 1.53991494948552447182;
         Q[2] := 0.982403709157920235114;
         Q[3] := 0.325732924782444448493;
         Q[4] := 0.0563921837420478160373;
         Q[5] := 0.00410369723978904575884;

         Result := Y + EvaluatePolynomial6(P, z - 1.5) / EvaluatePolynomial6(Q, z - 1.5);
         Result := Result * (exp(-z * z) / z);
      end
      else if (z < 4.5) then begin
         // Maximum Deviation Found:                     1.512e-17
         // Expected Error Term:                         1.512e-17
         // Maximum Relative Change in Control Points:   2.222e-04
         // Max Error found at double precision =        2.062515e-17
         Y := 0.5405750274658203125;

         P[0] := 0.00295276716530971662634;
         P[1] := 0.0137384425896355332126;
         P[2] := 0.00840807615555585383007;
         P[3] := 0.00212825620914618649141;
         P[4] := 0.000250269961544794627958;
         P[5] := 0.113212406648847561139e-4;

         Q[0] := 1;
         Q[1] := 1.04217814166938418171;
         Q[2] := 0.442597659481563127003;
         Q[3] := 0.0958492726301061423444;
         Q[4] := 0.0105982906484876531489;
         Q[5] := 0.000479411269521714493907;

         Result := Y + EvaluatePolynomial6(P, z - 3.5) / EvaluatePolynomial6(Q, z - 3.5);
         Result := Result * (exp(-z * z) / z);
      end
      else begin
         // Max Error found at double precision =        2.997958e-17
         // Maximum Deviation Found:                     2.860e-17
         // Expected Error Term:                         2.859e-17
         // Maximum Relative Change in Control Points:   1.357e-05
         Y := 0.5579090118408203125;

         P[0] := 0.00628057170626964891937;
         P[1] := 0.0175389834052493308818;
         P[2] := -0.212652252872804219852;
         P[3] := -0.687717681153649930619;
         P[4] := -2.5518551727311523996;
         P[5] := -3.22729451764143718517;
         P[6] := -2.8175401114513378771;

         Q[0] := 1;
         Q[1] := 2.79257750980575282228;
         Q[2] := 11.0567237927800161565;
         Q[3] := 15.930646027911794143;
         Q[4] := 22.9367376522880577224;
         Q[5] := 13.5064170191802889145;
         Q[6] := 5.48409182238641741584;

         Result := Y + EvaluatePolynomial7(P, 1 / z) / EvaluatePolynomial7(Q, 1 / z);
         Result := Result * (exp(-z * z) / z);
      end
   end
   else
   begin
      //
      // Any value of z larger than 28 will underflow to zero:
      //
      Result := 0;
      invert := not invert;
   end;

   if (invert) then
      Result := 1 - Result;
end;

function erfc(z: double): double;
begin
  Result := ErfImp(z,True);
end;

function expm1(x: double): double;
var
  a: double;
  Y: double;
  n: array[0..6] of double;
  d: array[0..6] of double;
begin
   a := Abs(x);
   if (a > 0.5) then begin
      if (a >= LogMaxValue) then begin
         if (x > 0) then begin
            Result := expm1(0);
            Exit;
         end;
         Result := -1;
         Exit;
      end;
      Result := Exp(x) - 1;
      Exit;
   end;
   if (a < const_epsilon) then begin
      Result := x;
      Exit;
   end;

   Y := 0.10281276702880859375e1;

   n[0] := -0.281276702880859375e-1;
   n[1] := 0.512980290285154286358e0;
   n[2] := -0.667758794592881019644e-1;
   n[3] := 0.131432469658444745835e-1;
   n[4] := -0.72303795326880286965e-3;
   n[5] := 0.447441185192951335042e-4;
   n[6] := -0.714539134024984593011e-6;

   d[0] := 1;
   d[1] := -0.461477618025562520389e0;
   d[2] := 0.961237488025708540713e-1;
   d[3] := -0.116483957658204450739e-1;
   d[4] := 0.873308008461557544458e-3;
   d[5] := -0.387922804997682392562e-4;
   d[6] := 0.807473180049193557294e-6;

   Result := x * Y + x * EvaluatePolynomial6(n, x) / EvaluatePolynomial6(d, x);
end;

function log1p(x: double): double;
begin
  Result := LnXp1(x);
end;

function powm1(a,z: double): double; 
var
  p: double;
begin
   if ((Abs(a) < 1) or (Abs(z) < 1)) then begin
      p := Ln(a) * z;
      if (Abs(p) < 2) then begin
         Result := expm1(p);
         Exit;
      end;
      // otherwise fall though:
   end;
   Result := Power(a, z) - 1;
end;

function itrunc(V: double): integer;
begin
  Result := Trunc(V);
end;

end.
