unit XLSMath5;

interface

uses Classes, SysUtils, Math,
     XLSUtils5, XLSMathData5;


function BetaCDF(a, b, x: double; Inverse: boolean = False): double;
function BetaPDF(a, b, x: double): double;
function BetaInvCDF(a, b, x: double): double;
//function BetaInv(a, b, p: double): double;
function HypergeometricDistCDF(x, r, n, N_: longword): double;
function HypergeometricDistPDF(x, r, n, N_: longword): double;
function HypergeometricDistCDFImp(x, r, n, N_: longword; invert: boolean): double;
function HypergeometricDistPDFImp(x, r, n, N_: longword): double;
function BinomialCDF(k: longword; n: longword; p: double): double;
function BinomialCDF2(k: longword; n: longword; p: double): double;
function BinomialPDF(k: longword; n: longword; p: double): double;
function BinomialQuantile(n: integer; p,p2: double): double;
function ChiSquaredDistributionCDF(DegreesOfFreedom,ChiSquare: double): double;
function ChiSquaredDistributionPDF(DegreesOfFreedom,ChiSquare: double): double;
function ChiSquaredDistributionInvPDF(df,x: double): double;
function GaussInvPDF(x,mean: double): double;
function ExponentialDistCDF(x,lambda: double): double;
function ExponentialDistPDF(x,lambda: double): double;
function GammaCDF(x,shape,scale: double): double;
function GammaPDF(x,shape,scale: double): double;
function GammaDistributionQuantile(a,b,p: double): double;
function NormalCDF(x, mean, stdev: double): double;
function NormalPDF(x, mean, stdev: double): double;
function NormalQuantile(p, mean, stdev: double): double;
function LogNormalCDF(x, mean, stdev: double): double;
function LogNormalPDF(x, mean, stdev: double): double;
function LogNormalCDFInv(x, mean, stdev: double): double;
function LogNormalQuantile(p, mean, stdev: double): double;
function NegBinomialCDF(k, r, p: double): double;
function NegBinomialPDF(k, r, p: double): double;
function PoissonCDF(k, mean: double): double;
function PoissonPDF(k, mean: double): double;
function StudentTDistCDF(t ,DegreesOfFreedom: double): double;
function StudentTDistPDF(t ,DegreesOfFreedom: double): double;
function WeibullCDF(x, shape,scale: double): double;
function WeibullPDF(x, shape,scale: double): double;
function InverseStudentsT(df, u, v: double; pexact: PBoolean = Nil): double;
function ChiSquaredQuantile(DegreesOfFreedom, p: double): double;
function FisherFDistCDF(x, df1, df2: double): double;
function FisherFDistPDF(x, df1, df2: double): double;
function FisherFDistQuantile(p, df1, df2: double): double;

type TFunc2DoubleBool = function(a,b: double): boolean;

type TFloatPair = record
     a,b: double;
     end;

type TFloatTriple = record
     a,b,c: double;
     end;

type TIterateFunc = class(TObject)
protected
public
     function Iterate: double; overload; virtual; abstract;
     function Iterate(x: double): double; overload; virtual; abstract;
     end;

type TIterateFuncTriple = class(TObject)
protected
public
     function Iterate(AValue: double): TFloatTriple; virtual; abstract;
     end;

type TIterateFuncPair = class(TObject)
protected
public
     function Iterate: TFloatPair; overload; virtual; abstract;
     function Iterate(x: double): TFloatPair; overload; virtual; abstract;
     end;

type TIBetaSeriesIter = class(TIterateFunc)
protected
     FRes: double;
     Fx: double;
     Fapn: double;
     Fpoch: double;
     Fn: integer;
public
     constructor Create(a,b,x,mult: double);

     function Iterate: double; override;
     end;

type TLowerIncompleteGammaSeriesIter = class(TIterateFunc)
protected
     Fa: double;
     Fz: double;
     FRes: double;
public
     constructor Create(a,z: double);

     function Iterate: double; override;
     end;

type TSmallGamma2SeriesIter = class(TIterateFunc)
protected
     Fa: double;
     Fx: double;
     Fapn: double;
     Fn: integer;
     FRes: double;
public
     constructor Create(a,x: double);

     function Iterate: double; override;
     end;

type TErfAsymptSeriesIter = class(TIterateFunc)
protected
     FRes: double;
     Fxx: double;
     Ftk: integer;
public
     constructor Create(z: double);

     function Iterate: double; override;
     end;

type DistributionQuantileIter = class(TIterateFunc)
protected
     FComp: boolean;
     FTarget: double;
     Fn: double;
     Fsuccess_fraction: double;
public
     constructor Create(n: longword; success_fraction, p,q: double);

     function Iterate(x: double): double; override;
     end;

type TIBetaFraction2Iter = class(TIterateFuncPair)
protected
     Fa: double;
     Fb: double;
     Fx: double;
     Fm: integer;
public
     constructor Create(a,b,x: double);

     function Iterate: TFloatPair; override;
     end;

type TUpperIncompleteGammaFractIter = class(TIterateFuncPair)
protected
     Fa: double;
     Fz: double;
     Fk: integer;
public
     constructor Create(a,z: double);

     function Iterate: TFloatPair; override;
     end;

type TTemmeRootFinderIter = class(TIterateFuncPair)
protected
     Ft: double;
     Fa: double;
public
     constructor Create(At,Aa: double);

     function Iterate(x: double): TFloatPair; override;
     end;

type TGammaPInverseIter = class(TIterateFuncTriple)
protected
     Fa: double;
     Fp: double;
     FInv: boolean;
public
     constructor Create(Aa,Ap: double; AInv: boolean);

     function Iterate(x: double): TFloatTriple; override;
     end;

type TIBetaRootsIter = class(TIterateFuncTriple)
protected
     Fa: double;
     Fb: double;
     FTarget: double;
     FInvert: boolean;
public
     constructor Create(Aa,Ab,At: double; AInverse: boolean = False);

     function Iterate(x: double): TFloatTriple; override;
     end;

type TXLSLanczos = class(TObject)
protected
public
     function g: double; virtual; abstract;
     function Sum(z: double): double; virtual; abstract;
     function SumExpGScaled(Z: double): double; virtual; abstract;
     end;

type TXLSLanczos13m53 = class(TXLSLanczos)
protected
public
     function g: double; override;
     function Sum(z: double): double; override;
     function SumExpGScaled(z: double): double; override;
     end;

function sinpx(z: double): double;

function EvaluateRational(Num: array of double; Denom: array of longword; Count: integer; z: double): double;
function SumSeries(Func: TIterateFunc; Bits, MaxTerms: integer; InitValue: double): double;
function Beta(A, B: double): double;
function IBetaSeries(a,b,x,s0: double; Lanczos: TXLSLanczos; Normalized: boolean; Derivative: PDouble; y: double): double;
function log1pmx(x: double): double; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
function RegularisedGammaPrefix(a,z: double; Lanczos: TXLSLanczos): double;
function TGammaDeltaRatioImpLanczos(z,delta: double; Lanczos: TXLSLanczos): double;
function TGammaDeltaRatioImp(z,delta: double): double;
function FiniteGammaQ(a,x: double): double;
function GammaImp(z: double; Lanczos: TXLSLanczos): double;
function FiniteHalfGammaQ(a,x: double; Derivative: PDouble): double;
function FullIGammaPrefix(a,z: double): double;
function LowerGammaSeries(a,z: double): double;
function TGamma(z: double): double;
function LGammaSmallImp(z,zm1,zm2: double): double;
function TGammap1m1Imp(dz: double): double;
function TGamma1pm1(z: double): double;
function TGammaSmallUpperPart(a,x: double): double;
function GammaIncompleteImp(a,x: double; normalized,invert: boolean; Derivative: PDouble): double;
function BetaSmallBLargeASeries(a,b,x,y,s0,mult: double; Normalized: boolean): double;
function InverseDiscreteQuantile(n: longword; success_fraction,p,q,guess,multiplier,adder: double; max_iter: integer): double;
function DoInverseDiscreteQuantile(n: longword; k,p,q,guess,multiplier,adder: double; TolFunc: TFunc2DoubleBool; var max_iter: integer; min_bound, max_bound: integer): double;
function RisingFactorialRatio(A,B,_P1: double): double;
function BinomialCCDF(n, k, x, y: double): double;
function BinomialQuantileImp(trials: integer; success_fraction,p,q: double): double;
function IBetaFraction2(a, b, x, y: double; Normalized: boolean; Derivative: PDouble): double;
function IBetaPowerTerms(a, b, x, y: double; Lcz: TXLSLanczos; Normalized: boolean): double;
function IbetaAStep(a,b,x,y: double; k: integer; Normalized: boolean; Derivative: PDouble): double;
function IBetaImp(a, b, x: double; Inv, Normalized: boolean; Derivative: PDouble): double;
function IBetaC(a, b, x: double): double;
function IBetaInvImp(a, b, p, q: double; py: PDouble): double;
function IBetaInv(a, b, p: double): double;
function IBeta(a, b, x: double): double;
function IGammaTemmeLarge(a,x: double): double;
function Erfc(z: double): double;
function erf_inv_imp(p, q: double): double;
function erfc_inv(z: double): double;
function ErfImp(z: double; Invert: boolean): double;
function Gamma_Q(a,z: double): double;
function Gamma_P(a,z: double): double;
function ContinuedFractionA(g: TIterateFuncPair; bits: integer): double;
function ContinuedFractionB(g: TIterateFuncPair; bits: integer): double;
function IBetaDerivativeImp(a, b, x: double): double;
function GammaPDerivativeImp(a, x: double): double;
function InverseStudentsTBodySeries(df, u: double): double;
function InverseStudentsTTailSeries(df, v: double): double;
function InverseStudentsTHill(ndf, u: double): double;
function GammaPInvImp(a, p: double): double;
function FindInverseGamma(a, p, q: double; has_10_digits: PBoolean): double;
function FindInverseS(p, q: double): double;
function LGammaImp(z: double; Lanczos: TXLSLanczos; sign: PInteger = Nil): double;
function DiDonato_SN(a, x: double; N: longword; tolerance: double = 0): double;
function HalleyIterate(Iter: TIterateFuncTriple; guess, min_, max_: double; digits: integer; var max_iter: integer): double;
procedure HandleZeroDerivative2(Iter: TIterateFuncPair; var last_f0, f0, delta, res, guess, min, max: double);
procedure HandleZeroDerivative3(Iter: TIterateFuncTriple; var last_f0, f0, delta, res, guess, min, max: double);
function FindIBetaInvFromTDist(a, p, q: double; py: PDouble): double;
function TemmeMethod1IbetaInverse(a, b, z: double): double;
function TemmeMethod2IBetaInverse(a, b, z, r, theta: double): double;
function TemmeMethod3IBetaInverse(a, b, p, q: double): double;
function SecantInterpolate(a, b, fa, fb: double): double;
function CubicInterpolate(a, b, d, e, fa, fb, fd, fe: double): double;
procedure Bracket(Iter: TIterateFunc; var a, b, c, fa, fb, d, fd: double);
function Toms748Solve(Iter: TIterateFunc; ax, bx, fax, fbx: double; TolFunc: TFunc2DoubleBool; var max_iter: integer): TFloatPair;
function NewtonRaphsonIterate(Iter: TIterateFuncPair; guess, min, max: double; digits: integer; var max_iter: integer): double;
function InverseBinomialCornishFisher(n, sf, p, q: double): double;

function HypergeometricPdfFactorialImp(x, r, n, N_: longword): double;
function HypergeometricPdfPrimeImp(x, r, n, N_: longword): double;
function HypergeometricPdfLanczosImp(x, r, n, N_: longword): double;

implementation


const
  const_e = 2.7182818284590452353602874713526624977572470936999595749669676;
  const_euler = 0.577215664901532860606512090082402431042159335939923598805;
  const_root_pi = 1.7724538509055160272981674833411451827975;
  const_root_two = 1.414213562373095048801688724209698078569671875376948073;
  MaxSeriesIterations = 100000;
  ResNumDigits = 53;
  LogMaxValue = 709;
  LogMinValue = -707;
  NiceArraySize = 30;


var
  Def_Lanczos: TXLSLanczos13m53;

// *****************************************************************************

procedure Swap(var A,B: double);
var
  T: double;
begin
  T := A;
  A := B;
  B := T;
end;

function EqualCeil(a,b: double): boolean;
begin
  Result := Ceil(a) = Ceil(b);
end;

function EqualFloor(a,b: double): boolean;
begin
  Result := Floor(a) = Floor(b);
end;

// Be carefule when using. Will evaluate both ATrue and AFalse.
function IIF(ACondition: boolean; ATrue,AFalse: double): double; overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
begin
  if ACondition then
    Result := ATrue
  else
    Result := AFalse;
end;

function IIF(ACondition: boolean; ATrue,AFalse: boolean): boolean; overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
begin
  if ACondition then
    Result := ATrue
  else
    Result := AFalse;
end;

// *****************************************************************************

function sinpx(z: double): double;
var
  sign: integer;
  fl: double;
  dist: double;
begin
   // Ad hoc function calculates x * sin(pi * x),
   // taking extra care near when x is near a whole number.

   sign := 1;
   if z < 0 then
      z := -z
   else
      sign := -sign;
   fl := Floor(z);
   if Odd(Trunc(fl)) then begin
//      fl := fl + 1;
      dist := fl - z;
      sign := -sign;
   end
   else
      dist := z - fl;

   if dist > 0.5 then
      dist := 1 - dist;
   Result := dist * Sin(Pi);
   Result := sign * z * result;
end;

function EvaluateRational(Num: array of double; Denom: array of longword; Count: integer; z: double): double;
var
  i: integer;
  s1,s2: double;
begin
  if z <= 1 then begin
    s1 := Num[Count - 1];
    s2 := Denom[Count - 1];
    for i := Count - 2 downto 0 do begin
      s1 := s1 * z;
      s2 := s2 * z;
      s1 := s1 + num[i];
      s2 := s2 + denom[i];
    end
  end
  else begin
    z := 1 / z;
    s1 := Num[0];
    s2 := Denom[0];
    for i := 1 to count - 1 do begin
      s1 := s1 * z;
      s2 := s2 * z;
      s1 := s1 + num[i];
      s2 := s2 + denom[i];
    end;
  end;
  Result := s1 / s2;
end;

function SumSeries(Func: TIterateFunc; Bits, MaxTerms: integer; InitValue: double): double;
var
  Counter: integer;
  Factor: double;
  NextTerm: double;
begin
   Counter := MaxTerms;

   Factor := Ldexp(1,Bits);
   Result := InitValue;
   repeat
     NextTerm := Func.Iterate;
     Result := Result + NextTerm;
     Dec(Counter);
   until not ((Abs(Result) < Abs(Factor * NextTerm)) and (Counter > 0));

   // set max_terms to the actual number of terms of the series evaluated:
//   MaxTerms := MaxTerms - Counter;
end;

// *****************************************************************************

function BetaImp(a, b: double; Lanczos: TXLSLanczos): double;
var
  prefix: double;
  c: double;
  agh: double;
  bgh: double;
  cgh: double;
  ambh: double;
begin
   prefix := 1;
   c := a + b;

   // Special cases:
   if ((c = a) and (b < const_epsilon)) then begin
      Result := TGamma(b);
      Exit;
   end
   else if ((c = b) and (a < const_epsilon)) then begin
      Result := TGamma(a);
      Exit;
   end;
   if (b = 1) then begin
      Result := 1 / a;
      Exit;
   end
   else if(a = 1) then begin
      Result := 1 / b;
      Exit;
   end;

//   /*
//   //
//   // This code appears to be no longer necessary: it was
//   // used to offset errors introduced from the Lanczos
//   // approximation, but the current Lanczos approximations
//   // are sufficiently accurate for all z that we can ditch
//   // this.  It remains in the file for future reference...
//   //
//   // If a or b are less than 1, shift to greater than 1:
//   if(a < 1)
//   begin
//      prefix *= c / a;
//      c += 1;
//      a += 1;
//   end
//   if(b < 1)
//   begin
//      prefix *= c / b;
//      c += 1;
//      b += 1;
//   end
//   */

   if (a < b) then
      Swap(a, b);

   // Lanczos calculation:
   agh := a + Lanczos.g() - 0.5;
   bgh := b + Lanczos.g() - 0.5;
   cgh := c + Lanczos.g() - 0.5;
   result := Lanczos.SumExpGScaled(a) * Lanczos.SumExpGScaled(b) / Lanczos.SumExpGScaled(c);
   ambh := a - 0.5 - b;
   if ((Abs(b * ambh) < (cgh * 100)) and (a > 100)) then
      // Special case where the base of the power term is close to 1
      // compute (1+x)^y instead:
      result := result * exp(ambh * log1p(-b / cgh))
   else
      result := result * Power(agh / cgh, a - 0.5 - b);
   if (cgh > 1e10) then
      // this avoids possible overflow, but appears to be marginally less accurate:
      result := result * Power((agh / cgh) * (bgh / cgh), b)
   else
      result := result * Power((agh * bgh) / (cgh * cgh), b);
   result := result * sqrt(const_e / bgh);

   // If a and b were originally less than 1 we need to scale the result:
   result := result * prefix;
end;

function Beta(a, b: double): double;
begin
  Result := BetaImp(a,b,Def_Lanczos);
end;

function IBetaSeries(a,b,x,s0: double; Lanczos: TXLSLanczos; Normalized: boolean; Derivative: PDouble; y: double): double;
var
  c: double;
  agh: double;
  bgh: double;
  cgh: double;
  s: TIBetaSeriesIter;
begin
  if Normalized then begin
    c := a + b;

    // incomplete beta power term, combined with the Lanczos approximation:
    agh := a + Lanczos.g - 0.5;
    bgh := b + Lanczos.g - 0.5;
    cgh := c + Lanczos.g - 0.5;
    Result := Lanczos.SumExpGScaled(c) / (Lanczos.SumExpGScaled(a) * Lanczos.SumExpGScaled(b));
    if a * b < bgh * 10 then
//         Result := Result * Exp((b - 0.5) * boost::math::log1p(a / bgh, pol));
      Result := Result * Exp((b - 0.5) * log1p(a / bgh))
    else
      Result := Result * Power(cgh / bgh, b - 0.5);
    Result := Result * Power(x * cgh / agh, a);
    Result := Result * Sqrt(agh / const_e);

    if Derivative <> Nil then
      Derivative^ := Result * Power(y, b);
  end
  else
    // Non-normalised, just compute the power:
    Result := Power(x, a);
  if Result < MinDouble then begin
     Result := s0; // Safeguard: series can't cope with denorms.
     Exit;
  end;
  s := TIBetaSeriesIter.Create(a,b,x,Result);
  try
    Result := SumSeries(s, ResNumDigits, MaxSeriesIterations, s0);
  finally
    s.Free;
  end;
end;

function log1pmx(x: double): double; 
begin
  Result := Ln(1 + x) - x;
end;

function RegularisedGammaPrefix(a,z: double; Lanczos: TXLSLanczos): double;
var
  agh: double;
  alz: double;
  amz: double;
  amza: double;
  sq: double;
  prefix: double;
  d: double;
begin
   agh := a + Lanczos.g() - 0.5;
   d := ((z - a) - Lanczos.g() + 0.5) / agh;

//   if a < 1 then begin
//      //
//      // We have to treat a < 1 as a special case because our Lanczos
//      // approximations are optimised against the factorials with a > 1,
//      // and for high precision types especially (128-bit reals for example)
//      // very small values of a can give rather eroneous results for gamma
//      // unless we do this:
//      //
//      // TODO: is this still required?  Lanczos approx should be better now?
//      //
//      // Axolot: Lanczos has definitley improved since Jobbik entered the Hungarian parlament.
//      if z <= LogMinValue then
//      begin
//         // Oh dear, have to use logs, should be free of cancellation errors though:
////         Result := Exp(a * Ln(z) - z - lgamma_imp(a, pol, l));
//      end
//      else
//      begin
//         // direct calculation, no danger of overflow as gamma(a) < 1/a
//         // for small a.
//         Result := Power(z, a) * Exp(-z) / gamma_imp(a, pol, l);
//      end
//   end
   {else} if (Abs(d*d*a) <= 100) and (a > 150) then begin
      // special case for large a and a ~ z.
      prefix := a * log1pmx(d) + z * (0.5 - Lanczos.g()) / agh;
      prefix := Exp(prefix);
   end
   else
   begin
      //
      // general case.
      // direct computation is most accurate, but use various fallbacks
      // for different parts of the problem domain:
      //
      alz := a * Ln(z / agh);
      amz := a - z;
      if (Min(alz, amz) <= LogMinValue) or (Max(alz, amz) >= LogMaxValue) then begin
         amza := amz / a;
         if (Min(alz, amz)/2 > LogMinValue) and (Max(alz, amz)/2 < LogMaxValue) then begin
            // compute square root of the result and then square it:
            sq := Power(z / agh, a / 2) * Exp(amz / 2);
            prefix := sq * sq;
         end
         else if (Min(alz, amz)/4 > LogMinValue) and (Max(alz, amz)/4 < LogMaxValue) and (z > a) then begin
            // compute the 4th root of the result then square it twice:
            sq := Power(z / agh, a / 4) * exp(amz / 4);
            prefix := sq * sq;
            prefix := prefix * prefix;
         end
         else if (amza > LogMinValue) and (amza < LogMaxValue) then
            prefix := Power((z * exp(amza)) / agh, a)
         else
            prefix := Exp(alz + amz);
      end
      else
         prefix := Power(z / agh, a) * Exp(amz);
   end;
   prefix := prefix * sqrt(agh / const_e) / Lanczos.SumExpGScaled(a);
   Result := prefix;
end;

function TGammaDeltaRatioImpLanczos(z,delta: double; Lanczos: TXLSLanczos): double;
var
  zgh: double;
begin
   zgh := z + Lanczos.g() - 0.5;

   if Abs(delta) < 10 then
      result := Exp((0.5 - z) * log1p(delta / zgh))
   else
      result := Power(zgh / (zgh + delta), z - 0.5);
   Result := Result * (Power(const_e / (zgh + delta), delta));
   Result := Result * (Lanczos.Sum(z) / Lanczos.Sum(z + delta));
end;

function TGammaDeltaRatioImp(z,delta: double): double;
begin
   if Floor(delta) = delta then begin
      if Floor(z) = z then begin
         //
         // Both z and delta are integers, see if we can just use table lookup
         // of the factorials to get the result:
         //
         if ((z <= MAX_QUICKFACTORIAL) and (z + delta <= MAX_QUICKFACTORIAL)) then begin
           Result := QuickFactorial(Trunc(z) - 1) / QuickFactorial(Trunc(z + delta) - 1);
           Exit;
         end;
      end;
      if Abs(delta) < 20 then begin
         //
         // delta is a small integer, we can use a finite product:
         //
         if delta = 0 then begin
            Result := 1;
            Exit;
         end;
         if Delta < 0 then begin
            z := z - 1;
            Result := z;
            while 0 <> Trunc(delta) do begin
               z := z - 1;
               Result := Result * z;
               delta := delta + 1;
            end;
            Exit;
         end
         else
         begin
            Result := 1 / z;
            delta := delta - 1;
            while 0 <> Trunc(delta) do begin
               z := z + 1;
               Result := Result / z;
               delta := delta - 1;
            end;
            Exit;
         end
      end
   end;
   Result := TGammaDeltaRatioImpLanczos(z, delta, Def_Lanczos);
end;

function FiniteGammaQ(a,x: double): double;
var
  n: integer;
  sum: double;
  term: double;
begin
   //
   // Calculates normalised Q when a is an integer:
   //
   sum := Exp(-x);
   if sum <> 0 then begin
      term := sum;
      n := 1;
      while n < a do begin
         term := term / n;
         term := term * x;
         sum := sum + term;
         Inc(n);
      end;
   end;
   Result := sum;
end;

function GammaImp(z: double; Lanczos: TXLSLanczos): double;
var
  zgh: double;
  hp: double;
begin
   Result := 1;

   if z <= 0 then begin
      if Floor(z) = z then
         raise XLSRWException.Create('Evaluation  of tgamma at a negative integer');
      if z <= -20 then begin
         result := GammaImp(-z, Lanczos) * sinpx(z);
         if((Abs(result) < 1) and (MaxDouble * Abs(result) < Pi)) then
            raise XLSRWException.Create('Result of tgamma is too large to represent.');
         Result := -Pi / result;
         if result = 0 then
            raise XLSRWException.Create('Result of tgamma is too small to represent.');
         Exit;
      end;

      // shift z to > 1:
      while z < 0 do begin
         Result := Result / z;
         z := z + 1;
      end;
   end;
   if (Floor(z) = z) and (z <= MAX_QUICKFACTORIAL) then
      Result := Result * QuickFactorial(Trunc(z) - 1)
   else
   begin
      Result := Result * Lanczos.Sum(z);
      if (z * Ln(z) > LogMaxValue) then begin
         // we're going to overflow unless this is done with care:
         zgh := (z + Lanczos.g()) - 0.5;
         if (Ln(zgh) * z / 2 > LogMaxValue) then
            raise XLSRWException.Create('Result of tgamma is too large to represent.');
         hp := Power(zgh, (z / 2) - 0.25);
         Result := Result * (hp / Exp(zgh));
         if (MaxDouble / hp < Result) then
            raise XLSRWException.Create('Result of tgamma is too large to represent.');
         Result := Result * hp;
      end
      else begin
         zgh := (z + Lanczos.g()) - 0.5;
         Result := Result * (Power(zgh, z - 0.5) / Exp(zgh));
      end;
   end;
end;

function FiniteHalfGammaQ(a,x: double; Derivative: PDouble): double;
var
  n: integer;
  e: double;
  term: double;
  sum: double;
  half: double;
begin
   //
   // Calculates normalised Q when a is a half-integer:
   //
   e := ErfImp(Sqrt(x),True);
   if ((e <> 0) and (a > 1)) then begin
      term := Exp(-x) / Sqrt(Pi * x);
      term := term * x;
      half := 1 / 2;
      term := term / half;
      sum := term;
      n := 2;
      while n < a do begin
         term := term / (n - half);
         term := term * x;
         sum := sum + term;
         Inc(n);
      end;
      e := e + sum;
      if (Derivative <> Nil) then
         Derivative^ := 0;
   end
   else if (Derivative <> Nil) then
      // We'll be dividing by x later, socalculate derivative * x:
      Derivative^ := sqrt(x) * exp(-x) / const_root_pi;
   Result := e;
end;

function FullIGammaPrefix(a,z: double): double;
var
  prefix: double;
  alz: double;
begin
   alz := a * Ln(z);

   if (z >= 1)  then begin
      if((alz < LogMaxValue) and (-z > LogMinValue)) then
         prefix := Power(z, a) * exp(-z)
      else if(a >= 1) then
         prefix := Power(z / exp(z/a), a)
      else
         prefix := exp(alz - z);
   end
   else
   begin
      if (alz > LogMinValue) then
         prefix := Power(z, a) * exp(-z)
      else if (z/a < LogMaxValue) then
         prefix := Power(z / exp(z/a), a)
      else
         prefix := exp(alz - z);
   end;
   //
   // This error handling isn't very good: it happens after the fact
   // rather than before it...
   //
//   if((boost::math::fpclassify)(prefix) == (int)FP_INFINITE)
//      policies::raise_overflow_error<T>("boost::math::detail::full_igamma_prefix<%1%>(%1%, %1%)", "Result of incomplete gamma function is too large to represent.", pol);

   Result := prefix;
end;

function LowerGammaSeries(a,z: double): double;
var
  Iter: TLowerIncompleteGammaSeriesIter;
begin
  Iter := TLowerIncompleteGammaSeriesIter.Create(a,z);
  try
    Result := SumSeries(Iter,53,12,0);
  finally
    Iter.Free;
  end;
end;

function TGamma(z: double): double;
begin
  Result := GammaImp(z,Def_Lanczos);
end;

function LGammaSmallImp(z,zm1,zm2: double): double;
var
  r: double;
  R_: double;
  prefix: double;
  Y: double;
  P: array[0..7] of double;
  Q: array[0..7] of double;
begin
   // This version uses rational approximations for small
   // values of z accurate enough for 64-bit mantissas
   // (80-bit long doubles), works well for 53-bit doubles as well.
   // Lanczos is only used to select the Lanczos function.

   Result := 0;
   if (z < const_epsilon) then begin
      Result := -Ln(z);
      Exit;
   end
   else if ((zm1 = 0) or (zm2 = 0)) then begin
      // nothing to do, result is zero....
   end
   else if (z > 2) then begin
      // Begin by performing argument reduction until
      // z is in [2,3):
      if (z >= 3) then begin
         repeat
            z := z - 1;
            zm2 := zm2 - 1;
            Result := Result + ln(z);
         until not (z >= 3);
         // Update zm2, we need it below:
         zm2 := z - 2;
      end;

      //
      // Use the following form:
      //
      // lgamma(z) = (z-2)(z+1)(Y + R(z-2))
      //
      // where R(z-2) is a rational approximation optimised for
      // low absolute error - as long as it's absolute error
      // is small compared to the constant Y - then any rounding
      // error in it's computation will get wiped out.
      //
      // R(z-2) has the following properties:
      //
      // At double: Max error found:                    4.231e-18
      // At long double: Max error found:               1.987e-21
      // Maximum Deviation Found (approximation error): 5.900e-24
      //
      P[0] := -0.180355685678449379109e-1;
      P[1] := 0.25126649619989678683e-1;
      P[2] := 0.494103151567532234274e-1;
      P[3] := 0.172491608709613993966e-1;
      P[4] := -0.259453563205438108893e-3;
      P[5] := -0.541009869215204396339e-3;
      P[6] := -0.324588649825948492091e-4;

      Q[0] := 0.1e1;
      Q[1] := 0.196202987197795200688e1;
      Q[2] := 0.148019669424231326694e1;
      Q[3] := 0.541391432071720958364e0;
      Q[4] := 0.988504251128010129477e-1;
      Q[5] := 0.82130967464889339326e-2;
      Q[6] := 0.224936291922115757597e-3;
      Q[7] := -0.223352763208617092964e-6;

      Y := 0.158963680267333984375e0;

      r := zm2 * (z + 1);
      R_ := EvaluatePolynomial7(P, zm2);
      R_ := R_ / EvaluatePolynomial8(Q, zm2);

      Result := Result + r * Y + r * R_;
   end
   else
   begin
      //
      // If z is less than 1 use recurrance to shift to
      // z in the interval [1,2]:
      //
      if (z < 1) then begin
         Result := Result + -Ln(z);
         zm2 := zm1;
         zm1 := z;
         z := z + 1;
      end;
      //
      // Two approximations, on for z in [1,1.5] and
      // one for z in [1.5,2]:
      //
      if (z <= 1.5) then begin
         //
         // Use the following form:
         //
         // lgamma(z) = (z-1)(z-2)(Y + R(z-1))
         //
         // where R(z-1) is a rational approximation optimised for
         // low absolute error - as long as it's absolute error
         // is small compared to the constant Y - then any rounding
         // error in it's computation will get wiped out.
         //
         // R(z-1) has the following properties:
         //
         // At double precision: Max error found:                1.230011e-17
         // At 80-bit long double precision:   Max error found:  5.631355e-21
         // Maximum Deviation Found:                             3.139e-021
         // Expected Error Term:                                 3.139e-021

         //
         Y := 0.52815341949462890625;

         P[0] := 0.490622454069039543534e-1;
         P[1] := -0.969117530159521214579e-1;
         P[2] := -0.414983358359495381969e0;
         P[3] := -0.406567124211938417342e0;
         P[4] := -0.158413586390692192217e0;
         P[5] := -0.240149820648571559892e-1;
         P[6] := -0.100346687696279557415e-2;


         Q[0] := 0.1e1;
         Q[1] := 0.302349829846463038743e1;
         Q[2] := 0.348739585360723852576e1;
         Q[3] := 0.191415588274426679201e1;
         Q[4] := 0.507137738614363510846e0;
         Q[5] := 0.577039722690451849648e-1;
         Q[6] := 0.195768102601107189171e-2;

         r := EvaluatePolynomial7(P, zm1) / EvaluatePolynomial7(Q, zm1);
         prefix := zm1 * zm2;

         Result := Result + prefix * Y + prefix * r;
      end
      else begin
         //
         // Use the following form:
         //
         // lgamma(z) = (2-z)(1-z)(Y + R(2-z))
         //
         // where R(2-z) is a rational approximation optimised for
         // low absolute error - as long as it's absolute error
         // is small compared to the constant Y - then any rounding
         // error in it's computation will get wiped out.
         //
         // R(2-z) has the following properties:
         //
         // At double precision, max error found:              1.797565e-17
         // At 80-bit long double precision, max error found:  9.306419e-21
         // Maximum Deviation Found:                           2.151e-021
         // Expected Error Term:                               2.150e-021
         //
         Y := 0.452017307281494140625;

         P[0] := -0.292329721830270012337e-1;
         P[1] := 0.144216267757192309184e0;
         P[2] := -0.142440390738631274135e0;
         P[3] := 0.542809694055053558157e-1;
         P[4] := -0.850535976868336437746e-2;
         P[5] := 0.431171342679297331241e-3;

         Q[0] := 0.1e1;
         Q[1] := -0.150169356054485044494e1;
         Q[2] := 0.846973248876495016101e0;
         Q[3] := -0.220095151814995745555e0;
         Q[4] := 0.25582797155975869989e-1;
         Q[5] := -0.100666795539143372762e-2;
         Q[6] := -0.827193521891290553639e-6;

         r := zm2 * zm1;
         R_ := EvaluatePolynomial6(P, -zm2) / EvaluatePolynomial7(Q, -zm2);

         Result := Result + r * Y + r * R_;
      end;
   end;
end;

function TGammap1m1Imp(dz: double): double;
begin
   if (dz < 0) then begin
	  if (dz < -0.5) then
		 // Best method is simply to subtract 1 from tgamma:
		 Result := TGamma(1 + dz) - 1
	  else
		 // Use expm1 on lgamma:
		 Result := expm1(-log1p(dz) + LGammaSmallImp(dz+2, dz + 1, dz));
   end
   else begin
	  if (dz < 2) then
		 // Use expm1 on lgamma:
		 Result := expm1(LGammaSmallImp(dz+1, dz, dz-1))
	  else
		 // Best method is simply to subtract 1 from tgamma:
		 Result := TGamma(1 + dz) - 1;
   end;
end;

function TGamma1pm1(z: double): double;
begin
  Result := TGammap1m1Imp(z);
end;

function TGammaSmallUpperPart(a,x: double): double;
var
  zero: double;
  s: TSmallGamma2SeriesIter;
begin
   Result := TGamma1pm1(a) - powm1(x, a);
   Result := Result / a;
   s := TSmallGamma2SeriesIter.Create(a,x);
   try
     zero := 0;
     Result := Result - Power(x, a) * SumSeries(s,ResNumDigits,MaxSeriesIterations,zero);
   finally
     s.Free;
   end;
end;

function ContinuedFractionA(g: TIterateFuncPair; bits: integer): double;
var
  factor: double;
  tiny: double;
  v: TFloatPair;
  f: double;
  C: double;
  D: double;
  delta: double;
  a0: double;
begin
   factor := ldexp(1.0, 1-bits); // 1 / pow(result_type(2), bits);
   tiny := MinDouble;

   v := g.Iterate;

   f := v.b;
   a0 := v.a;
   if (f = 0) then
      f := tiny;
   C := f;
   D := 0;

   repeat
      v := g.Iterate;
      D := v.b + v.a * D;
      if (D = 0) then
         D := tiny;
      C := v.b + v.a / C;
      if (C = 0) then
         C := tiny;
      D := 1/D;
      delta := C*D;
      f := f * delta;
   until not (Abs(delta - 1) > factor);

   Result := a0/f;
end;

function UpperGammaFraction(a, z: double; bits: integer): double;
var
  f: TUpperIncompleteGammaFractIter;
begin
   // Multiply result by z^a * e^-z to get the full
   // upper incomplete integral.  Divide by tgamma(z)
   // to normalise.
   f := TUpperIncompleteGammaFractIter.Create(a,z);
   try
     Result := 1 / (z - a + 1 + ContinuedFractionA(f, bits));
   finally
     f.Free;
   end;
end;

function ErfImp(z: double; Invert: boolean): double;
var
  x: double;
  s: TErfAsymptSeriesIter;
begin
   if (z < 0) then begin
      if (not invert) then
         Result := -ErfImp(-z, invert)
      else
         Result := 1 + ErfImp(-z, false);
      Exit;
   end;

   if (not invert and (z > MaxSingle)) then begin
      s := TErfAsymptSeriesIter.Create(z);
      try
        Result := SumSeries(s,ResNumDigits, MaxSeriesIterations, 1);
      finally
        s.Free;
      end;
   end
   else begin
      x := z * z;
      if (x < 0.6) then begin
         // Compute P:
         result := z * exp(-x);
         result := result / sqrt(Pi);
         if (result <> 0) then
            result := result * 2 * LowerGammaSeries(0.5, x);
      end
      else if (x < 1.1) then begin
         // Compute Q:
         invert := not invert;
         result := TGammaSmallUpperPart(0.5, x);
         result := result / sqrt(Pi);
      end
      else
      begin
         // Compute Q:
         invert := not invert;
         result := z * exp(-x);
         result := result / sqrt(Pi);
         result := result * UpperGammaFraction(0.5, x, ResNumDigits);
      end;
   end;
   if (invert) then
      result := 1 - result;
end;

function Erfc(z: double): double;
begin
  Result := ErfImp(z,True);
end;

function IGammaTemmeLarge(a,x: double): double;
const C0: array[0..6] of double = (
   -0.333333333,
    0.0833333333,
   -0.0148148148,
    0.00115740741,
    0.000352733686,
   -0.000178755144,
    0.391926318e-4);

const C1: array[0..4] of double = (
   -0.00185185185,
   -0.00347222222,
    0.00264550265,
   -0.000990226337,
    0.000205761317);

const C2: array[0..2] of double = (
    0.00413359788,
   -0.00268132716,
    0.000771604938);
var
  sigma: double;
  phi: double;
  y: double;
  z: double;
  workspace: array[0..2] of double;
begin
   sigma := (x - a) / a;
   phi := -log1pmx(sigma);
   y := a * phi;
   z := sqrt(2 * phi);
   if (x < a) then
      z := -z;

   workspace[0] := EvaluatePolynomial7(C0, z);

   workspace[1] := EvaluatePolynomial5(C1, z);

   workspace[2] := EvaluatePolynomial3(C2, z);

   Result := EvaluatePolynomial3(workspace, 1 / a);
   Result := Result * (exp(-y) / sqrt(2 * Pi * a));
   if (x < a)  then
      Result := -Result;

   Result := Result + Erfc(sqrt(y)) / 2;
end;

function GammaIncompleteImp(a,x: double; normalized,invert: boolean; Derivative: PDouble): double;
var
  is_int: boolean;
  is_half_int: boolean;
  is_small_a: boolean;
  use_temme: boolean;
  sigma: double;
  gam: double;
begin
   is_int := floor(a) = a;
   is_half_int := (floor(2 * a) = 2 * a) and not is_int;
   is_small_a := (a < 30) and (a <= x + 1);

   if (is_int and is_small_a and (x > 0.6)) then begin
      // calculate Q via finite sum:
      invert := not invert;
      result := FiniteGammaQ(a, x);
      if not normalized then
         Result := Result * GammaImp(a,Def_Lanczos);
      // TODO: calculate derivative inside sum:
      if (Derivative <> Nil) then
         Derivative^ := RegularisedGammaPrefix(a, x, Def_Lanczos);
   end
   else if (is_half_int and is_small_a and (x > 0.2)) then begin
      // calculate Q via finite sum for half integer a:
      invert := not invert;
      result := FiniteHalfGammaQ(a, x, Derivative);
      if (normalized = false) then
         Result := Result * GammaImp(a,Def_Lanczos);
      if(Derivative <> Nil) and (Derivative^ = 0) then
         Derivative^ := RegularisedGammaPrefix(a, x, Def_Lanczos);
   end
   else if (x < 0.5) then begin
      //
      // Changeover criterion chosen to give a changeover at Q ~ 0.33
      //
      if (-0.4 / Ln(x) < a) then begin
         // Compute P:
         if normalized then
           Result := RegularisedGammaPrefix(a, x, Def_Lanczos)
         else
           Result := FullIGammaPrefix(a, x);
         if (Derivative <> Nil) then
            Derivative^ := Result;
         if (Result <> 0) then
            Result := Result * (LowerGammaSeries(a, x) / a);
      end
      else
      begin
         // Compute Q:
         invert := not invert;
         result := TGammaSmallUpperPart(a, x);
         if (normalized) then
            Result := Result / TGamma(a);
         if (Derivative <> Nil) then
            Derivative^ := RegularisedGammaPrefix(a, x, Def_Lanczos);
      end
   end
   else if (x < 1.1) then begin
      //
      // Changover here occurs when P ~ 0.6 or Q ~ 0.4:
      //
      if (x * 1.1 < a) then begin
         // Compute P:
         if normalized then
           Result := RegularisedGammaPrefix(a, x, Def_Lanczos)
         else
           Result := FullIGammaPrefix(a, x);
         if (Derivative <> Nil) then
            Derivative^ := Result;
         if (Result <> 0) then
            Result := Result * (LowerGammaSeries(a, x) / a);
      end
      else
      begin
         // Compute Q:
         invert := not invert;
         result := TGammaSmallUpperPart(a, x);
         if (normalized) then
            Result := Result / TGamma(a);
         if (Derivative <> Nil)then
            Derivative^ := RegularisedGammaPrefix(a, x, Def_Lanczos);
      end;
   end
   else begin
      //
      // Begin by testing whether we're in the "bad" zone
      // where the result will be near 0.5 and the usual
      // series and continued fractions are slow to converge:
      //
      use_temme := false;
      if (normalized and IsSpecialized and (a > 20)) then begin
         sigma := Abs((x-a)/a);
         if ((a > 200) and (ResNumDigits <= 113)) then begin
            //
            // This limit is chosen so that we use Temme's expansion
            // only if the result would be larger than about 10^-6.
            // Below that the regular series and continued fractions
            // converge OK, and if we use Temme's method we get increasing
            // errors from the dominant erfc term as it's (inexact) argument
            // increases in magnitude.
            //
            if (20 / a > sigma * sigma) then
               use_temme := true;
         end
         else if (ResNumDigits <= 64) then begin
            // Note in this zone we can't use Temme's expansion for
            // types longer than an 80-bit real:
            // it would require too many terms in the polynomials.
            if (sigma < 0.4) then
               use_temme := true;
         end;
      end;
      if (use_temme) then begin
         //
         // Use compile time dispatch to the appropriate
         // Temme asymptotic expansion.  This may be dead code
         // if T does not have numeric limits support, or has
         // too many digits for the most precise version of
         // these expansions, in that case we'll be calling
         // an empty function.
         //
         result := IGammaTemmeLarge(a, x);
         if (x >= a) then
            invert := not invert;
         if (Derivative <> Nil) then
            Derivative^ := RegularisedGammaPrefix(a, x, Def_Lanczos);
      end
      else
      begin
         //
         // Regular case where the result will not be too close to 0.5.
         //
         // Changeover here occurs at P ~ Q ~ 0.5
         //
         if normalized then
           Result := RegularisedGammaPrefix(a, x, Def_Lanczos)
         else
           Result := FullIGammaPrefix(a, x);
         if (Derivative <> Nil) then
            Derivative^ := result;
         if (x < a) then begin
            // Compute P:
            if (result <> 0) then
               result := result * (LowerGammaSeries(a, x) / a);
         end
         else
         begin
            // Compute Q:
            invert := not invert;
            if (result <> 0) then
               result := result * UpperGammaFraction(a, x, ResNumDigits);
         end
      end
   end;

   if (invert) then begin
      if normalized then
        gam := 1
      else
        gam := TGamma(a);
      result := gam - result;
   end;
   if (Derivative <> Nil) then begin
      //
      // Need to convert prefix term to derivative:
      //
      if((x < 1) and (MaxDouble * x < Derivative^)) then
         // overflow, just return an arbitrarily large value:
         Derivative^ := MaxDouble / 2
      else
        Derivative^ := Derivative^ / x;
   end;
end;

function Gamma_Q(a,z: double): double;
begin
  Result := GammaIncompleteImp(a,z,True,True,Nil);
end;

function Gamma_P(a,z: double): double;
begin
  Result := GammaIncompleteImp(a,z,True,False,Nil);
end;

function BetaSmallBLargeASeries(a,b,x,y,s0,mult: double; Normalized: boolean): double;
var
  bm1: double;
  t: double;
  lx: double;
  u: double;
  prefix: double;
  h: double;
  j: double;
  sum: double;
  lx2: double;
  lxp: double;
  t4: double;
  b2n: double;
  mbn: double;
  r: double;
  n: integer;
  m: integer;
  tnp1: longword;
  tmp1: longword;
  p: array[0..NiceArraySize - 1] of double;
begin
   // This is DiDonato and Morris's BGRAT routine, see Eq's 9 through 9.6.
   //
   // Some values we'll need later, these are Eq 9.1:
   //
   bm1 := b - 1;
   t := a + bm1 / 2;
   if y < 0.35 then
      lx := log1p(-y)
   else
      lx := ln(x);
   u := -t * lx;
   // and from from 9.2:
   h := RegularisedGammaPrefix(b, u, Def_Lanczos);
   if h <= MinDouble then begin
      Result := s0;
      Exit;
   end;
   if normalized then
   begin
      prefix := h / TGammaDeltaRatioImp(a, b);
      prefix := prefix / Power(t, b);
   end
   else
   begin
     raise XLSRWException.Create('only normalized supported');
//      prefix := full_igamma_prefix(b, u, pol) / pow(t, b);
   end;
   prefix := prefix * mult;
   //
   // now we need the quantity Pn, unfortunatately this is computed
   // recursively, and requires a full history of all the previous values
   // so no choice but to declare a big table and hope it's big enough...
   //
//   T p[ ::boost::math::detail::Pn_size<T>::value ] = { 1 };  // see 9.3.
   //
   // Now an initial value for J, see 9.6:
   //
   j := Gamma_Q(b, u) / h;
   //
   // Now we can start to pull things together and evaluate the sum in Eq 9:
   //
   sum := s0 + prefix * j;  // Value at N = 0
   // some variables we'll need:
   tnp1 := 1; // 2*N+1
   lx2 := lx / 2;
   lx2 := lx2 * lx2;
   lxp := 1;
   t4 := 4 * t * t;
   b2n := b;

   for n := 1 to High(p) do begin
//      // debugging code, enable this if you want to determine whether
//      // the table of Pn's is large enough...
//      //
//      static int max_count := 2;
//      if(n > max_count)
//      begin
//         max_count := n;
//         std::cerr << "Max iterations in BGRAT was " << n << std::endl;
//      end
      //
      // begin by evaluating the next Pn from Eq 9.4:
      //
      tnp1 := tnp1 + 2;
      p[n] := 0;
//      mbn := b - n;  // "Value assigned to 'mbn' never used. Checked and true.
      tmp1 := 3;
      for m := 1 to n - 1 do begin
         mbn := m * b - n;
         p[n] := p[n] + mbn * p[n-m] / QuickFactorial(tmp1);
         tmp1 := tmp1 + 2;
      end;
      p[n] := p[n] / n;
      p[n] := p[n] + bm1 / QuickFactorial(tnp1);
      //
      // Now we want Jn from Jn-1 using Eq 9.6:
      //
      j := (b2n * (b2n + 1) * j + (u + b2n + 1) * lxp) / t4;
      lxp := lxp * lx2;
      lxp := lxp + 2;
      //
      // pull it together with Eq 9:
      //
      r := prefix * p[n] * j;
      sum := sum + r;
      if(r > 1) then begin
         if (Abs(r) < Abs(const_epsilon * sum)) then
            break;
      end
      else
      begin
         if (Abs(r / const_epsilon) < Abs(sum)) then
            break;
      end;
   end;
   Result := sum;
end;

function RisingFactorialRatio(A,B,_P1: double): double;
begin
  raise XLSRWException.Create('TODO RisingFactorialRatio');
end;

function BinomialCCDF(n, k, x, y: double): double;
var
  i: integer;
  Term: double;
begin
  Result := Power(x, n);
  Term := Result;
  i := Trunc(n - 1);
  while i > k do begin
    Term := Term * (((i + 1) * y) / ((n - i) * x));
    Result := Result + Term;
    Dec(i);
  end;
end;

function SecantInterpolate(a, b, fa, fb: double): double;
var
  c,tol: double;
begin
   //
   // Performs standard secant interpolation of [a,b] given
   // function evaluations f(a) and f(b).  Performs a bisection
   // if secant interpolation would leave us very close to either
   // a or b.  Rationale: we only call this function when at least
   // one other form of interpolation has already failed, so we know
   // that the function is unlikely to be smooth with a root very
   // close to a or b.
   //
   tol := const_epsilon * 5;
   c := a - (fa / (fb - fa)) * (b - a);
   if ((c <= a + Abs(a) * tol) or (c >= b - Abs(b) * tol)) then
     Result := (a + b) / 2
   else
     Result := c;
end;

procedure Bracket(Iter: TIterateFunc; var a, b, c, fa, fb, d, fd: double);
var
  tol,fc: double;
begin
   //
   // Given a point c inside the existing enclosing interval
   // [a, b] sets a = c if f(c) == 0, otherwise finds the new
   // enclosing interval: either [a, c] or [c, b] and sets
   // d and fd to the point that has just been removed from
   // the interval.  In other words d is the third best guess
   // to the root.
   //
   tol := const_epsilon * 2;
   //
   // If the interval [a,b] is very small, or if c is too close
   // to one end of the interval then we need to adjust the
   // location of c accordingly:
   //
   if ((b - a) < 2 * tol * a) then
      c := a + (b - a) / 2
   else if (c <= a + Abs(a) * tol) then
      c := a + Abs(a) * tol
   else if(c >= b - Abs(b) * tol) then
      c := b - Abs(a) * tol;
   //
   // OK, lets invoke f(c):
   //
   fc := Iter.Iterate(c);
   //
   // if we have a zero then we have an exact solution to the root:
   //
   if (fc = 0) then begin
      a := c;
      fa := 0;
      d := 0;
      fd := 0;
      Exit;
   end;
   //
   // Non-zero fc, update the interval:
   //
   if (Sign(fa) * Sign(fc) < 0) then begin
      d := b;
      fd := fb;
      b := c;
      fb := fc;
   end
   else begin
      d := a;
      fd := fa;
      a := c;
      fa := fc;
   end;
end;

function QuadraticInterpolate(a, b, d, fa, fb, fd: double; count: integer): double;
var
  i: integer;
  A_,B_,c: double;

function SafeDiv(num, denom, r: double): double;
begin
   //
   // return num / denom without overflow,
   // return r if overflow would occur.
   //
   if (Abs(denom) < 1) then begin
      if(Abs(denom * MaxDouble) <= Abs(num)) then begin
         Result := r;
         Exit;
      end;
   end;
   Result := num / denom;
end;

begin
   //
   // Performs quadratic interpolation to determine the next point,
   // takes count Newton steps to find the location of the
   // quadratic polynomial.
   //
   // Point d must lie outside of the interval [a,b], it is the third
   // best approximation to the root, after a and b.
   //
   // Note: this does not guarentee to find a root
   // inside [a, b], so we fall back to a secant step should
   // the result be out of range.
   //
   // Start by obtaining the coefficients of the quadratic polynomial:
   //
   B_ := SafeDiv(fb - fa, b - a, MaxDouble);
   A_ := SafeDiv(fd - fb, d - b, MaxDouble);
   A_ := SafeDiv(A_ - B_, d - a, 0);

   if (a = 0) then begin
      // failure to determine coefficients, try a secant step:
      Result := SecantInterpolate(a, b, fa, fb);
      Exit;
   end;
   //
   // Determine the starting point of the Newton steps:
   //
   if (Sign(A_) * Sign(fa) > 0) then
      c := a
   else
      c := b;
   //
   // Take the Newton steps:
   //
   for i := 1 to count do
      c := c - SafeDiv(fa+(B_+A_*(c-b))*(c-a), (B_+ A_ * (2 * c - a - b)), 1 + c - a);
   if ((c <= a) or (c >= b)) then
      // Oops, failure, try a secant step:
      c := SecantInterpolate(a, b, fa, fb);
   Result := c;
end;

function CubicInterpolate(a, b, d, e, fa, fb, fd, fe: double): double;
var
  q11,q21,q31,d21,d31,q22,q32,d32,q33,c: double;
begin
   //
   // Uses inverse cubic interpolation of f(x) at points
   // [a,b,d,e] to obtain an approximate root of f(x).
   // Points d and e lie outside the interval [a,b]
   // and are the third and forth best approximations
   // to the root that we have found so far.
   //
   // Note: this does not guarentee to find a root
   // inside [a, b], so we fall back to quadratic
   // interpolation in case of an erroneous result.
   //
   q11 := (d - e) * fd / (fe - fd);
   q21 := (b - d) * fb / (fd - fb);
   q31 := (a - b) * fa / (fb - fa);
   d21 := (b - d) * fd / (fd - fb);
   d31 := (a - b) * fb / (fb - fa);
   q22 := (d21 - q11) * fb / (fe - fb);
   q32 := (d31 - q21) * fa / (fd - fa);
   d32 := (d31 - q21) * fd / (fd - fa);
   q33 := (d32 - q22) * fa / (fe - fa);
   c := q31 + q32 + q33 + a;

   if ((c <= a) or (c >= b)) then
      // Out of bounds step, fall back to quadratic interpolation:
      c := QuadraticInterpolate(a, b, d, fa, fb, fd, 3);

   Result := c;
end;

function Toms748Solve(Iter: TIterateFunc; ax, bx, fax, fbx: double; TolFunc: TFunc2DoubleBool; var max_iter: integer): TFloatPair;
var
   a,b,c,fa,fb,u,fu,a0,b0,d,fd,e,fe,mu: double;
   temp: double;
   min_diff: double;
   prof: boolean;
   count: integer;
begin
   //
   // Main entry point and logic for Toms Algorithm 748
   // root finder.
   //
   count := max_iter;
   mu := 0.5;

   // initialise a, b and fa, fb:
   a := ax;
   b := bx;
   if (a >= b) then
      raise XLSRWException.Create('Parameters a and b out of order');
   fa := fax;
   fb := fbx;

   if ((TolFunc(a,b)) or (fa = 0) or (fb = 0)) then begin
      max_iter := 0;
      if (fa = 0) then
         b := a
      else if (fb = 0) then
         a := b;
      Result.a := a;
      Result.b := b;
      Exit;
   end;

   if (Sign(fa) * Sign(fb) > 0) then
         raise XLSRWException.Create('Parameters a and b do not bracket the root.');
   // dummy value for fd, e and fe:
   fe := 1e5;
   e := 1e5;
   fd := 1e5;

   if (fa <> 0) then begin
      //
      // On the first step we take a secant step:
      //
      c := SecantInterpolate(a, b, fa, fb);
      Bracket(Iter, a, b, c, fa, fb, d, fd);
      Dec(count);

      if ((count > 0) and (fa <> 0) and not TolFunc(a, b)) then begin
         //
         // On the second step we take a quadratic interpolation:
         //
         c := QuadraticInterpolate(a, b, d, fa, fb, fd, 2);
         e := d;
         fe := fd;
         Bracket(Iter, a, b, c, fa, fb, d, fd);
         Dec(count);
      end;
   end;

   while ((count > 0) and (fa <> 0) and not TolFunc(a, b)) do begin
      // save our brackets:
      a0 := a;
      b0 := b;
      //
      // Starting with the third step taken
      // we can use either quadratic or cubic interpolation.
      // Cubic interpolation requires that all four function values
      // fa, fb, fd, and fe are distinct, should that not be the case
      // then variable prof will get set to true, and we'll end up
      // taking a quadratic step instead.
      //
      min_diff := MinDouble * 32;
      prof := (Abs(fa - fb) < min_diff) or (Abs(fa - fd) < min_diff) or (Abs(fa - fe) < min_diff) or (Abs(fb - fd) < min_diff) or (Abs(fb - fe) < min_diff) or (Abs(fd - fe) < min_diff);
      if (prof) then
         c := QuadraticInterpolate(a, b, d, fa, fb, fd, 2)
      else
         c := CubicInterpolate(a, b, d, e, fa, fb, fd, fe);
      //
      // re-bracket, and check for termination:
      //
      e := d;
      fe := fd;
      Bracket(Iter, a, b, c, fa, fb, d, fd);
      if( (0 = --count) or (fa = 0) or TolFunc(a, b)) then
         break;
      //
      // Now another interpolated step:
      //
      prof := (Abs(fa - fb) < min_diff) or (Abs(fa - fd) < min_diff) or (Abs(fa - fe) < min_diff) or (Abs(fb - fd) < min_diff) or (Abs(fb - fe) < min_diff) or (Abs(fd - fe) < min_diff);
      if (prof) then
         c := QuadraticInterpolate(a, b, d, fa, fb, fd, 3)
      else
         c := CubicInterpolate(a, b, d, e, fa, fb, fd, fe);
      //
      // Bracket again, and check termination condition, update e:
      //
      Bracket(Iter, a, b, c, fa, fb, d, fd);
      Dec(Count);
      if ((0 = count) or (fa = 0) or TolFunc(a, b)) then
         break;
      //
      // Now we take a double-length secant step:
      //
      if( Abs(fa) < Abs(fb)) then begin
         u := a;
         fu := fa;
      end
      else begin
         u := b;
         fu := fb;
      end;
      c := u - 2 * (fu / (fb - fa)) * (b - a);
      if (Abs(c - u) > (b - a) / 2) then
         c := a + (b - a) / 2;
      //
      // Bracket again, and check termination condition:
      //
      e := d;
      fe := fd;
      Bracket(Iter, a, b, c, fa, fb, d, fd);
      Dec(count);
      if((0 = count) or (fa = 0) or TolFunc(a, b)) then
         break;
      //
      // And finally... check to see if an additional bisection step is
      // to be taken, we do this if we're not converging fast enough:
      //
      if ((b - a) < mu * (b0 - a0)) then
         Continue;
      //
      // bracket again on a bisection:
      //
      e := d;
      fe := fd;
      temp := a + (b - a) / 2;
      Bracket(Iter, a, b, temp, fa, fb, d, fd);
      Dec(count);
   end;

   max_iter := max_iter - count;
   if (fa = 0) then
      b := a
   else if(fb = 0) then
      a := b;
   Result.a := a;
   Result.b := b;
end;

function DoInverseDiscreteQuantile(n: longword; k,p,q,guess,multiplier,adder: double; TolFunc: TFunc2DoubleBool; var max_iter: integer; min_bound, max_bound: integer): double;
var
  Iter: DistributionQuantileIter;
  a,b,fa,fb: double;
  r: TFloatPair;
  count: integer;
begin
//   distribution_quantile_finder<Dist> f(dist, p, q);
   //
   // Max bounds of the distribution:
   //
   if (guess > max_bound) then
      guess := max_bound;
   if (guess < min_bound) then
      guess := min_bound;

   Iter := DistributionQuantileIter.Create(n,k,p,q);
   try
     fa := Iter.Iterate(guess);
     count := max_iter - 1;
     fb := fa;
     a := guess;
     b :=0;

     if (fa = 0) then begin
        Result := guess;
        Exit;
     end;

     //
     // For small expected results, just use a linear search:
     //
     if (guess < 10) then begin
        b := a;
        while ((a < 10) and (fa * fb >= 0)) do begin
           if (fb <= 0) then begin
              a := b;
              b := a + 1;
              if (b > max_bound) then
                 b := max_bound;
              fb := Iter.Iterate(b);
              Dec(count);
              if (fb = 0) then begin
                 Result := b;
                 Exit;
              end;
           end
           else begin
              b := a;
              a := Max(b - 1, 0);
              if (a < min_bound) then
                 a := min_bound;
              fa := Iter.Iterate(a);
              Dec(count);
              if (fa = 0) then begin
                 Result := a;
                 Exit;
              end;
           end;
        end;
     end
     //
     // Try and bracket using a couple of additions first,
     // we're assuming that "guess" is likely to be accurate
     // to the nearest int or so:
     //
     else if (adder <> 0) then begin
        //
        // If we're looking for a large result, then bump "adder" up
        // by a bit to increase our chances of bracketing the root:
        //
        //adder := (std::max)(adder, 0.001f * guess);
        if (fa < 0) then begin
           b := a + adder;
           if (b > max_bound) then
              b := max_bound;
        end
        else begin
           b := Max(a - adder, 0);
           if (b < min_bound) then
              b := min_bound;
        end;
        fb := Iter.Iterate(b);
        Dec(count);
        if (fb = 0) then begin
           Result :=  b;
           Exit;
        end;
        if ((count > 0) and (fa * fb >= 0)) then begin
           //
           // We didn't bracket the root, try
           // once more:
           //
           a := b;
           fa := fb;
           if (fa < 0) then begin
              b := a + adder;
              if (b > max_bound) then
                 b := max_bound;
           end
           else begin
              b := Max(a - adder, 0);
              if (b < min_bound) then
                 b := min_bound;
           end;
           fb := Iter.Iterate(b);
           Dec(count);
        end;
        if (a > b) then begin
           swap(a, b);
           swap(fa, fb);
        end;
     end;
     //
     // If the root hasn't been bracketed yet, try again
     // using the multiplier this time:
     //
     if (Sign(fb) = Sign(fa)) then begin
        if (fa < 0) then begin
           //
           // Zero is to the right of x2, so walk upwards
           // until we find it:
           //
           while (Sign(fb) = Sign(fa)) do begin
              if (count = 0) then
                 raise XLSRWException.Create('Unable to bracket root');
              a := b;
              fa := fb;
              b := b * multiplier;
              if (b > max_bound) then
                 b := max_bound;
              fb := Iter.Iterate(b);
              Dec(count);
           end;
        end
        else begin
           //
           // Zero is to the left of a, so walk downwards
           // until we find it:
           //
           while (Sign(fb) = Sign(fa)) do begin
              if (Abs(a) < MinDouble) then begin
                 // Escape route just in case the answer is zero!
                 max_iter := max_iter - count;
                 max_iter := max_iter + 1;
                 Result := 0;
                 Exit;
              end;
              if (count = 0) then
                 raise XLSRWException.Create('Unable to bracket root');
              b := a;
              fb := fa;
              a := a / multiplier;
              if (a < min_bound) then
                 a := min_bound;
              fa := Iter.Iterate(a);
              Dec(count);
           end;
        end;
     end;
     max_iter := max_iter - count;
     if(fa = 0) then begin
        Result := a;
        Exit;
     end;
     if (fb = 0) then begin
        Result := b;
        Exit;
     end;
     //
     // Adjust bounds so that if we're looking for an integer
     // result, then both ends round the same way:
     //
     a := a + const_epsilon * a;
     //
     // We don't want zero or denorm lower bounds:
     //
     if (a < MinDouble) then
        a := MinDouble;
     //
     // Go ahead and find the root:
     //
     r := Toms748Solve(Iter, a, b, fa, fb, TolFunc,count);
   finally
     Iter.Free;
   end;
   max_iter := max_iter + max_iter + count;
   Result := (r.a + r.b) / 2;
end;

function InverseDiscreteQuantile(n: longword; success_fraction, p,q,guess,multiplier,adder: double; max_iter: integer): double;
begin
  if p < BinomialPDF(0,n,success_fraction) then
    Result := 0
  else if(p < 0.5) then
    Result := floor(DoInverseDiscreteQuantile(n,success_fraction,p,q,IIF(guess < 1,1,floor(guess)),multiplier,adder,EqualFloor,max_iter,0,100))
  else
    Result := ceil(DoInverseDiscreteQuantile(n,success_fraction, p,q,ceil(guess),multiplier,adder,EqualCeil,max_iter,0,100));
end;

function BinomialQuantileImp(trials: integer; success_fraction,p,q: double): double;
var
  guess,factor: double;
begin
    if (p = 0) then begin
      Result := 0;
      Exit;
    end;
    if (p = 1) then begin
      Result := trials;
      Exit;
    end;
    if (p <= Power(1 - success_fraction, trials)) then begin
      Result := 0;
      Exit;
    end;

    // Solve for quantile numerically:
    //
    guess := InverseBinomialCornishFisher(trials, success_fraction, p, q);
    factor := 8;
    if (trials > 100) then
       factor := 1.01 // guess is pretty accurate
    else if ((trials > 10) and (trials - 1 > guess) and (guess > 3)) then
       factor := 1.15 // less accurate but OK.
    else if (trials < 10) then begin
       // pretty inaccurate guess in this area:
       if (guess > trials / 64) then begin
          guess := trials / 4;
          factor := 2;
       end
       else
          guess := trials / 1024;
    end
    else
       factor := 2; // trials largish, but in far tails.

    Result := InverseDiscreteQuantile(trials,success_fraction,p,q,guess,factor,1,200);
end;

function ContinuedFractionB(g: TIterateFuncPair; bits: integer): double;
var
  factor: double;
  tiny: double;
  f: double;
  C: double;
  D: double;
  delta: double;
  v: TFloatPair;
begin
   factor := ldexp(1.0, 1 - bits); // 1 / pow(result_type(2), bits);
   tiny := MinDouble;

   v := g.Iterate;

   f := v.b;;
   if (f = 0) then
      f := tiny;
   C := f;
   D := 0;

   repeat
      v := g.Iterate;
      D := v.b + v.a * D;
      if (D = 0) then
         D := tiny;
      C := v.b + v.a / C;
      if (C = 0) then
         C := tiny;
      D := 1/D;
      delta := C*D;
      f := f * delta;
   until not (Abs(delta - 1) > factor);

   Result := f;
end;

function IBetaFraction2(A, B, X, Y: double; Normalized: boolean; Derivative: PDouble): double;
var
  fract: double;
  f: TIBetaFraction2Iter;
begin
   result := IBetaPowerTerms(a, b, x, y, Def_Lanczos, Normalized);
   if (Derivative <> Nil) then
     Derivative^ := result;
   if (result = 0) then
	   Exit;

   f := TIBetaFraction2Iter.Create(a,b,x);
   try
     fract := ContinuedFractionB(f, ResNumDigits);
   finally
     f.Free;
   end;
   result := result / fract;
end;

function IBetaPowerTerms(a, b, x, y: double; Lcz: TXLSLanczos; Normalized: boolean): double;
var
  Prefix: double;
  c: double;
  agh: double;
  bgh: double;
  cgh: double;
  l: double;
  l1: double;
  l2: double;
  l3: double;
  b1: double;
  b2: double;
  ratio: double;
  small_a: boolean;
begin
   if not Normalized then begin
     // can we do better here?
     Result := Power(x, a) * Power(y, b);
     Exit;
   end;

   Prefix := 1;
   c := a + b;

   // combine power terms with Lanczos approximation:
   agh := a + Lcz.g() - 0.5;
   bgh := b + Lcz.g() - 0.5;
   cgh := c + Lcz.g() - 0.5;
   Result := Lcz.SumExpGScaled(c) / (Lcz.SumExpGScaled(a) * Lcz.SumExpGScaled(b));

   // l1 and l2 are the base of the exponents minus one:
   l1 := (x * b - y * agh) / agh;
   l2 := (y * a - x * bgh) / bgh;
   if (Min(Abs(l1), Abs(l2)) < 0.2) then begin
      // when the base of the exponent is very near 1 we get really
      // gross errors unless extra care is taken:
      if (l1 * l2 > 0) or (Min(a, b) < 1) then begin
         //
         // This first branch handles the simple cases where either:
         //
         // * The two power terms both go in the same direction
         // (towards zero or towards infinity).  In this case if either
         // term overflows or underflows, then the product of the two must
         // do so also.
         // *Alternatively if one exponent is less than one, then we
         // can't productively use it to eliminate overflow or underflow
         // from the other term.  Problems with spurious overflow/underflow
         // can't be ruled out in this case, but it is *very* unlikely
         // since one of the power terms will evaluate to a number close to 1.
         //
         if Abs(l1) < 0.1 then
            Result := Result * Exp(a * log1p(l1))
         else
            Result := Result * Power((x * cgh) / agh, a);
         if Abs(l2) < 0.1 then
            Result := Result * Exp(b * log1p(l2))
         else
            Result := Result * Power((y * cgh) / bgh, b);
      end
      else if Max(Abs(l1), Abs(l2)) < 0.5 then begin
         //
         // Both exponents are near one and both the exponents are
         // greater than one and further these two
         // power terms tend in opposite directions (one towards zero,
         // the other towards infinity), so we have to combine the terms
         // to avoid any risk of overflow or underflow.
         //
         // We do this by moving one power term inside the other, we have:
         //
         //    (1 + l1)^a * (1 + l2)^b
         //  = ((1 + l1)*(1 + l2)^(b/a))^a
         //  = (1 + l1 + l3 + l1*l3)^a   ;  l3 = (1 + l2)^(b/a) - 1
         //                                    = exp((b/a) * log(1 + l2)) - 1
         //
         // The tricky bit is deciding which term to move inside :-)
         // By preference we move the larger term inside, so that the
         // size of the largest exponent is reduced.  However, that can
         // only be done as long as l3 (see above) is also small.
         //
         small_a := a < b;
         ratio := b / a;
         if (small_a and (ratio * l2 < 0.1)) or ( not small_a and (l1 / ratio > 0.1)) then begin
            l3 := Exp(ratio * log1p(l2)) - 1;
            l3 := l1 + l3 + l3 * l1;
            l3 := a * log1p(l3);
            Result := Result * Exp(l3);
         end
         else begin
            l3 := Exp(log1p(l1) / ratio) - 1;
            l3 := l2 + l3 + l3 * l2;
            l3 := b * log1p(l3);
            Result := Result * Exp(l3);
         end;
      end
      else if Abs(l1) < Abs(l2) then begin
         // First base near 1 only:
         l := a * log1p(l1) + b * Ln((y * cgh) / bgh);
         Result := Result * Exp(l);
      end
      else begin
         // Second base near 1 only:
         l := b * log1p(l2) + a * Ln((x * cgh) / agh);
         Result := Result * Exp(l);
      end;
   end
   else begin
      // general case:
      b1 := (x * cgh) / agh;
      b2 := (y * cgh) / bgh;
      l1 := a * Ln(b1);
      l2 := b * Ln(b2);
      if (l1 >= LogMaxValue) or (l1 <= LogMinValue) or (l2 >= LogMaxValue) or (l2 <= LogMinValue) then begin
         // Oops, overflow, sidestep:
         if a < b then
            Result := Result * Power(Power(b2, b/a) * b1, a)
         else
            Result := Result * Power(Power(b1, a/b) * b2, b);
      end
      else
         // finally the normal case:
         Result := Result * Power(b1, a) * Power(b2, b);
   end;
   // combine with the leftover terms from the Lanczos approximation:
   Result := Result * Sqrt(bgh / const_e);
   Result := Result * Sqrt(agh / cgh);
   Result := Result * prefix;
end;

function IbetaAStep(a,b,x,y: double; k: integer; Normalized: boolean; Derivative: PDouble): double;
var
  i: integer;
  prefix: double;
  sum: double;
  term: double;
begin
   prefix := IBetaPowerTerms(a, b, x, y, Def_Lanczos, Normalized);
   if Derivative <> Nil then
    Derivative^ := prefix;
//      BOOST_ASSERT(*p_derivative >= 0);
   prefix := prefix / a;
   if prefix = 0 then begin
      Result := prefix;
      Exit;
   end;
   sum := 1;
   term := 1;
   // series summation from 0 to k-1:
   for i := 0 to k - 2 do begin
      term := term * (a+b+i) * x / (a+i+1);
      sum := sum + term;
   end;
   prefix := prefix * sum;

   Result := prefix;
end;

function IBetaImp(a, b, x: double; Inv, Normalized: boolean; Derivative: PDouble): double;
var
  Invert: boolean;
  Fract: double;
  Y: double;
  Prefix: double;
  Lambda: double;
  K: double;
//  N: integer; ???
  N: double;
  BBar: double;
  Div_: double;
begin
  Invert := Inv;
  Y := 1 - X;

  if Derivative <> Nil then
    Derivative^ := -1; // value not set.

  if Normalized then begin
    if A = 0 then begin
      if B > 0 then begin
        Result := IIF(Inv,0,1);
        Exit;
      end;
    end
    else if B = 0 then begin
      if a > 0 then begin
        Result := IIF(Inv,1,0);
        Exit;
      end;
    end;
  end;

  if X = 0 then begin
    if Derivative <> Nil then
      Derivative^ := IIF(A = 1,1,IIF(A < 1,MaxDouble / 2,MinDouble * 2));
    if Invert then begin
      if Normalized then
        Result := 1
      else
        Result := Beta(a,b);
    end
    else
      Result := 0;
    Exit;
  end
  else if X = 1 then begin
    if Derivative <> Nil then
      Derivative^ := IIF(B = 1,1,IIF(B < 1,MaxDouble / 2,MinDouble * 2));

    if not Invert then begin
      if Normalized then
        Result := 1
      else
        Result := Beta(a,b);
    end
    else
      Result := 0;
    Exit;
  end;

  if Min(A,B) <= 1 then begin
    if X > 0.5 then begin
      Swap(A, B);
      Swap(X, Y);
      Invert := not Invert;
    end;
    if Max(A, B) <= 1 then begin
      // Both a,b < 1:
      if (A >= Min(0.2,B)) or (Power(X, A) <= 0.9) then begin
        if not Invert then
          Fract := IbetaSeries(A, B, X, 0, Def_Lanczos, Normalized, Derivative, y)
        else begin
          if Normalized then
            Fract := -1
          else
            Fract := -Beta(a,b);
          Invert := false;
          Fract := -IbetaSeries(A, B, X, Fract, Def_Lanczos, Normalized, Derivative, y);
        end
      end
      else begin
        Swap(A, B);
        Swap(X, Y);
        Invert := not Invert;
        if Y >= 0.3 then begin
          if not Invert then
            Fract := IbetaSeries(A, B, X, 0, Def_Lanczos, Normalized, Derivative, Y)
          else begin
            if Normalized then
              Fract := -1
            else
              Fract := -Beta(a,b);
            Invert := False;
            Fract := -IbetaSeries(A, B, X, Fract, Def_Lanczos, Normalized, Derivative, Y);
          end
        end
        else begin
          if not Normalized then
            Prefix := RisingFactorialRatio(A + B, A, 20)
          else
            Prefix := 1;
          Fract := IbetaAStep(A, B, X, Y, 20, Normalized, Derivative);
          if not Invert then
            Fract := BetaSmallBLargeASeries(A + 20, B, X, Y, Fract, Prefix, Normalized)
          else begin
            if Normalized then
              Fract := Fract - 1
            else
              Fract := Fract - Beta(a,b);
            Invert := False;
            fract := -BetaSmallBLargeASeries(A + 20, B, X, Y, Fract, Prefix, Normalized);
          end
        end
      end
    end
    else begin
      // One of a, b < 1 only:
      if (B <= 1) or ((X < 0.1) and (Power(B * X, A) <= 0.7)) then begin
        if not Invert then
          Fract := IbetaSeries(A, B, X, 0, Def_Lanczos, Normalized, Derivative, Y)
        else begin
          if Normalized then
            Fract := -1
          else
            Fract := -Beta(a,b);
          invert := false;
          fract := -IbetaSeries(A, B, X, Fract, Def_Lanczos, Normalized, Derivative, Y);
        end
      end
      else begin
        Swap(a, b);
        Swap(x, y);
        Invert := not Invert;
        if Y >= 0.3 then begin
          if not Invert then
            Fract := IbetaSeries(A, B, X, 0, Def_Lanczos, Normalized, Derivative, Y)
          else begin
            if Normalized then
              Fract := -1
            else
              Fract := -Beta(a,b);
            Invert := False;
            Fract := -IbetaSeries(A, B, X, Fract,Def_Lanczos, Normalized, Derivative, Y);
          end
        end
        else if A >= 15 then begin
          if not Invert then
            Fract := BetaSmallBLargeASeries(A, B, X, Y, 0, 1, Normalized)
          else begin
            if Normalized then
              Fract := -1
            else
              Fract := -Beta(a,b);
            Invert := False;
            Fract := -BetaSmallBLargeASeries(A, B, X, Y, Fract, 1, Normalized);
          end
        end
        else begin
          // Sidestep to improve errors:
          if not Normalized then
            Prefix := RisingFactorialRatio(A + B, A, 20)
          else
            Prefix := 1;
          Fract := IbetaAStep(A, B, X, Y, 20, Normalized, Derivative);
          if not Invert then
            Fract := BetaSmallBLargeASeries(A + 20, B, X, Y, Fract, Prefix, Normalized)
          else begin
            if Normalized then
              Fract := Fract - 1
            else
              Fract := Fract - Beta(a,b);
            Invert := False;
            fract := -BetaSmallBLargeASeries(A + 20, B, X, Y, Fract, Prefix, Normalized);
          end;
        end;
      end;
    end;
  end
  else begin
    // Both a,b >= 1:
    if a < b then
      Lambda := A - (A + B) * X
    else
      Lambda := (A + B) * Y - B;
    if Lambda < 0 then begin
      Swap(A, B);
      Swap(X, Y);
      invert := not invert;
    end;
    if B < 40 then begin
      if (Floor(A) = A) and (Floor(B) = B) then begin
        // relate to the binomial distribution and use a finite sum:
        K := A - 1;
        N := B + K;
        Fract := BinomialCCDF(N, K, X, Y);
        if not Normalized then
          Fract := Fract * Beta(A, B);
      end
      else if B * X <= 0.7 then begin
        if not Invert then
          Fract := IbetaSeries(A, B, X, 0, Def_Lanczos, Normalized, Derivative, Y)
        else begin
          if Normalized then
            Fract := -1
          else
            Fract := -Beta(a,b);
          Invert := False;
          Fract := -IbetaSeries(A, B, X, Fract, Def_Lanczos, Normalized, Derivative, Y);
        end
      end
      else if A > 15 then begin
        // sidestep so we can use the series representation:
        N := ITrunc(Floor(B));
        if N = B then
          N := N - 1;
        BBar := B - N;
        if not Normalized then
          Prefix := RisingFactorialRatio(A + BBar, BBar, N)
        else
          Prefix := 1;
        Fract := IbetaAStep(BBar, A, Y, X, Trunc(N), Normalized, Nil);
        Fract := BetaSmallBLargeASeries(A, BBar, X, Y, Fract, 1, Normalized);
        Fract := Fract / Prefix;
      end
      else if Normalized then begin
        // the formula here for the non-Normalized case is tricky to figure
        // out (for me!!), and requires two pochhammer calculations rather
        // than one, so leave it for now....
        N := ITrunc(Floor(B));
        BBar := B - N;
        if BBar <= 0 then begin
          N := N - 1;
          BBar := BBar + 1;
        end;
        Fract := IBetaAStep(BBar, A, Y, X, Trunc(N), Normalized, Nil);
        Fract := Fract + IBetaAStep(A, BBar, X, Y, 20, Normalized, Nil);
        if Invert then begin
          if Normalized then
            Fract := Fract - 1
          else
            Fract := Fract - Beta(a,b);
        end;
          //fract := ibeta_series(a+20, bbar, x, fract, l, Normalized, p_derivative, y);
        Fract := BetaSmallBLargeASeries(A + 20, BBar, X, Y, Fract, 1, Normalized);
        if Invert then begin
          Fract := -Fract;
          Invert := False;
        end
      end
      else
        Fract := IBetaFraction2(A, B, X, Y, Normalized, Derivative);
    end
    else
       Fract := IBetaFraction2(A, B, X, Y, Normalized, Derivative);
   end;
   if Derivative <> Nil then begin
     if Derivative^ < 0 then
       Derivative^ := IBetaPowerTerms(A, B, X, Y, Def_Lanczos, True);
     Div_ := Y * X;

     if Derivative <> Nil then begin
       if (MaxDouble * Div_ < Derivative^) then
         // overflow, return an arbitarily large value:
         Derivative^ := MaxDouble / 2
       else
        Derivative^ := Derivative^ / Div_;
     end;
   end;

   if Invert then begin
     if Normalized then
       Result := 1 - Fract
     else
       Result := Beta(A,B) - Fract;
   end
   else
     Result := Fract;

//  Result := IIF(Invert,IIF(Normalized,1,Beta(A,B)) - Fract,Fract);
end;

function FindIBetaInvFromTDist(a, p, q: double; py: PDouble): double;
var
  u,v,df,t,x: double;
begin
   u := IIF(p > q,(0.5 - q) / 2,p / 2);
   v := 1 - u; // u < 0.5 so no cancellation error
   df := a * 2;
   t := InverseStudentsT(df, u, v);
   x := df / (df + t * t);
   py^ := t * t / (df + t * t);
   Result := x;
end;

function GammaQInvImp(a, q: double): double;
var
  guess: double;
  has_10_digits: boolean;
  lower: double;
  digits: integer;
  max_iter: integer;
  Iter: TGammaPInverseIter;
begin
//   static const char* function = "boost::math::gamma_q_inv<%1%>(%1%, %1%)";

   if (a <= 0) then
      raise XLSRWException.Create('Argument a in the incomplete gamma function inverse must be >= 0.');
   if ((q < 0) or (q > 1)) then
      raise XLSRWException.Create('Probabilty must be in the range [0,1] in the incomplete gamma function inverse.');
   if (q = 0) then begin
      Result := MaxDouble;
      Exit;
   end;
   if (q = 1) then begin
      Result := 0;
      Exit;
   end;
   guess := FindInverseGamma(a, 1 - q, q, @has_10_digits);
   if ((ResNumDigits <= 36) and has_10_digits) then begin
      Result := guess;
      Exit;
   end;
   lower := MinDouble;
   if (guess <= lower) then
      guess := MinDouble;
   //
   // Work out how many digits to converge to, normally this is
   // 2/3 of the digits in T, but if the first derivative is very
   // large convergence is slow, so we'll bump it up to full
   // precision to prevent premature termination of the root-finding routine.
   //
   digits := ResNumDigits;
   if (digits < 30) then begin
      digits := digits * 2;
      digits := digits div 3;
   end
   else
   begin
      digits := digits div 2;
      digits := digits - 1;
   end;
   if ((a < 0.125) and (Abs(GammaPDerivativeImp(a, guess)) > 1 / sqrt(const_epsilon))) then
      digits := ResNumDigits;
   //
   // Go ahead and iterate:
   //
   max_iter := MaxSeriesIterations;
   Iter := TGammaPInverseIter.Create(a,q,True);
   try
     guess := HalleyIterate(Iter,guess,lower,MaxDouble,digits,max_iter);
   finally
     Iter.Free;
   end;
   if (guess = lower) then
      raise XLSRWException.Create('Expected result known to be non-zero, but is smaller than the smallest available number.');
   Result := guess;
end;

function TemmeMethod1IbetaInverse(a, b, z: double): double;
var
  c,x: double;
  r2: double;
  eta,eta0,eta_2: double;
  B_,B_2,B_3: double;
  terms: array[0..3] of double;
  workspace: array[0..6] of double;
begin
   r2 := sqrt(2);
   //
   // get the first approximation for eta from the inverse
   // error function (Eq: 2.9 and 2.10).
   //
   eta0 := erfc_inv(2 * z);
   eta0 := eta0 / -sqrt(a / 2);

   terms[0] := eta0;
   //
   // calculate powers:
   //
   B_ := b - a;
   B_2 := B_ * B_;
   B_3 := B_2 * B_;
   //
   // Calculate correction terms:
   //

   // See eq following 2.15:
   workspace[0] := -B_ * r2 / 2;
   workspace[1] := (1 - 2 * B_) / 8;
   workspace[2] := -(B_ * r2 / 48);
   workspace[3] := -1 / 192;
   workspace[4] := -B_ * r2 / 3840;
   terms[1] := EvaluatePolynomial5(workspace, eta0);
   // Eq Following 2.17:
   workspace[0] := B_ * r2 * (3 * B_ - 2) / 12;
   workspace[1] := (20 * B_2 - 12 * B_ + 1) / 128;
   workspace[2] := B_ * r2 * (20 * B_ - 1) / 960;
   workspace[3] := (16 * B_2 + 30 * B_ - 15) / 4608;
   workspace[4] := B_ * r2 * (21 * B_ + 32) / 53760;
   workspace[5] := (-32 * B_2 + 63) / 368640;
   workspace[6] := -B_ * r2 * (120 * B_ + 17) / 25804480;
   terms[2] := EvaluatePolynomial7(workspace, eta0);
   // Eq Following 2.17:
   workspace[0] := B_ * r2 * (-75 * B_2 + 80 * B_ - 16) / 480;
   workspace[1] := (-1080 * B_3 + 868 * B_2 - 90 * B_ - 45) / 9216;
   workspace[2] := B_ * r2 * (-1190 * B_2 + 84 * B_ + 373) / 53760;
   workspace[3] := (-2240 * B_3 - 2508 * B_2 + 2100 * B_ - 165) / 368640;
   terms[3] := EvaluatePolynomial4(workspace, eta0);
   //
   // Bring them together to get a final estimate for eta:
   //
   eta := EvaluatePolynomial4(terms, 1/a);
   //
   // now we need to convert eta to x, by solving the appropriate
   // quadratic equation:
   //
   eta_2 := eta * eta;
   c := -exp(-eta_2 / 2);
   if (eta_2 = 0) then
      x := 0.5
   else
      x := (1 + eta * sqrt((1 + c) / eta_2)) / 2;

   Result := x;
end;

function NewtonRaphsonIterate(Iter: TIterateFuncPair; guess, min, max: double; digits: integer; var max_iter: integer): double;
var
  f0,f1,last_f0: double;
  factor,delta,delta1,delta2: double;
  count: integer;
  Res2: TFloatPair;
begin
   f0 := 0;
   last_f0 := 0;
   result := guess;

   factor := ldexp(1.0, 1 - digits);
   delta := 1;
   delta1 := MaxDouble;
//   delta2 := MaxDouble;

   count := max_iter;

   repeat
      last_f0 := f0;
      delta2 := delta1;
      delta1 := delta;
      Res2 := Iter.Iterate(result);
      f0 := Res2.a;
      f1 := Res2.b;
      if (0 = f0) then
         break;
      if (f1 = 0) then
         // Oops zero derivative!!!
         HandleZeroDerivative2(Iter, last_f0, f0, delta, result, guess, min, max)
      else
         delta := f0 / f1;
      if (Abs(delta * 2) > Abs(delta2)) then
         // last two steps haven't converged, try bisection:
         delta := IIf(delta > 0,(result - min) / 2,(result - max) / 2);
      guess := result;
      result := result - delta;
      if (result <= min) then begin
         delta := 0.5 * (guess - min);
         result := guess - delta;
         if ((result = min) or (result = max)) then
            break;
      end
      else if (result >= max) then begin
         delta := 0.5 * (guess - max);
         result := guess - delta;
         if ((result = min) or (result = max)) then
            break;
      end;
      // update brackets:
      if (delta > 0) then
         max := guess
      else
         min := guess;
      Dec(count);
   until not (((count > 0) and (Abs(result * factor) < Abs(delta))));

   max_iter := max_iter - count;
end;

function InverseBinomialCornishFisher(n, sf, p, q: double): double;
var
  m,sigma,sk,x,x2,w: double;
begin
    // mean:
    m := n * sf;
    // standard deviation:
    sigma := sqrt(n * sf * (1 - sf));
    // skewness
    sk := (1 - 2 * sf) / sigma;
    // kurtosis:
    // T k := (1 - 6 * sf * (1 - sf) ) / (n * sf * (1 - sf));
    // Get the inverse of a std normal distribution:
    x := erfc_inv(IIF(p > q,2 * q,2 * p)) * const_root_two;
    // Set the sign:
    if (p < 0.5) then
       x := -x;
    x2 := x * x;
    // w is correction term due to skewness
    w := x + sk * (x2 - 1) / 6;
//    /*
//    // Add on correction due to kurtosis.
//    // Disabled for now, seems to make things worse?
//    //
//    if(n >= 10)
//       w += k * x * (x2 - 3) / 24 + sk * sk * x * (2 * x2 - 5) / -36;
//       */
    w := m + sigma * w;
    if (w < MinDouble) then
       Result := sqrt(MinDouble)
    else if(w > n) then
       Result := n
    else
      Result := w;
end;

//function TemmeMethod2IBetaInverse(T /*a*/, T /*b*/, T z, T r, T theta, const Policy& pol)
function TemmeMethod2IBetaInverse(a, b, z, r, theta: double): double;
const
  co1 : array[0..2] of double = ( -1, -5, 5);
  co2 : array[0..3] of double = ( 1, 21, -69, 46);
  co3 : array[0..4] of double = ( 7, -2, 33, -62, 31);
  co4 : array[0..5] of double = ( 25, -52, -17, 88, -115, 46);
  co5 : array[0..3] of double = ( 7, 12, -78, 52);
  co6 : array[0..4] of double = ( -7, 2, 183, -370, 185);
  co7 : array[0..5] of double = ( -533, 776, -1835, 10240, -13525, 5410);
  co8 : array[0..6] of double = ( -1579, 3747, -3372, -15821, 45588, -45213, 15071);
  co9 : array[0..5] of double = (449, -1259, -769, 6686, -9260, 3704);
  co10: array[0..6] of double = ( 63149, -151557, 140052, -727469, 2239932, -2251437, 750479);
  co11: array[0..7] of double = ( 29233, -78755, 105222, 146879, -1602610, 3195183, -2554139, 729754);
  co12: array[0..2] of double = ( 1, -13, 13);
  co13: array[0..3] of double = ( 1, 21, -69, 46);
var
  eta0: double;
  s,c,u,x: double;
  c_2,s_2,lu: double;
  sc,sc_2,sc_3,sc_4,sc_5,sc_6,sc_7: double;
  alpha,eta: double;
  lower, upper: double;
  nIter: integer;
  terms: array[0..3] of double;
  workspace: array[0..5] of double;
  Iter: TTemmeRootFinderIter;
begin
   eta0 := erfc_inv(2 * z);
   eta0 := eta0 / -sqrt(r / 2);

   s := sin(theta);
   c := cos(theta);
   //
   // Now we need to purturb eta0 to get eta, which we do by
   // evaluating the polynomial in 1/r at the bottom of page 151,
   // to do this we first need the error terms e1, e2 e3
   // which we'll fill into the array "terms".  Since these
   // terms are themselves polynomials, we'll need another
   // array "workspace" to calculate those...
   //
   terms[0] := eta0;
   //
   // some powers of sin(theta)cos(theta) that we'll need later:
   //
   sc := s * c;
   sc_2 := sc * sc;
   sc_3 := sc_2 * sc;
   sc_4 := sc_2 * sc_2;
   sc_5 := sc_2 * sc_3;
   sc_6 := sc_3 * sc_3;
   sc_7 := sc_4 * sc_3;
   //
   // Calculate e1 and put it in terms[1], see the middle of page 151:
   //
   workspace[0] := (2 * s * s - 1) / (3 * s * c);
   workspace[1] := -EvaluateEvenPolynomial3(co1, s) / (36 * sc_2);
   workspace[2] := EvaluateEvenPolynomial4(co2, s) / (1620 * sc_3);
   workspace[3] := -EvaluateEvenPolynomial5(co3, s) / (6480 * sc_4);
   workspace[4] := EvaluateEvenPolynomial6(co4, s) / (90720 * sc_5);
   terms[1] := EvaluatePolynomial5(workspace, eta0);
   //
   // Now evaluate e2 and put it in terms[2]:
   //
   workspace[0] := -EvaluateEvenPolynomial4(co5, s) / (405 * sc_3);
   workspace[1] := EvaluateEvenPolynomial5(co6, s) / (2592 * sc_4);
   workspace[2] := -EvaluateEvenPolynomial6(co7, s) / (204120 * sc_5);
   workspace[3] := -EvaluateEvenPolynomial7(co8, s) / (2099520 * sc_6);
   terms[2] := EvaluatePolynomial4(workspace, eta0);
   //
   // And e3, and put it in terms[3]:
   //
   workspace[0] := EvaluateEvenPolynomial6(co9, s) / (102060 * sc_5);
   workspace[1] := -EvaluateEvenPolynomial7(co10, s) / (20995200 * sc_6);
   workspace[2] := EvaluateEvenPolynomial8(co11, s) / (36741600 * sc_7);
   terms[3] := EvaluatePolynomial3(workspace, eta0);
   //
   // Bring the correction terms together to evaluate eta,
   // this is the last equation on page 151:
   //
   eta := EvaluatePolynomial4(terms, 1/r);
   //
   // Now that we have eta we need to back solve for x,
   // we seek the value of x that gives eta in Eq 3.2.
   // The two methods used are described in section 5.
   //
   // Begin by defining a few variables we'll need later:
   //
   s_2 := s * s;
   c_2 := c * c;
   alpha := c / s;
   alpha := alpha * alpha;
   lu := (-(eta * eta) / (2 * s_2) + ln(s_2) + c_2 * ln(c_2) / s_2);
   //
   // Temme doesn't specify what value to switch on here,
   // but this seems to work pretty well:
   //
   if (Abs(eta) < 0.7) then begin
      //
      // Small eta use the expansion Temme gives in the second equation
      // of section 5, it's a polynomial in eta:
      //
      workspace[0] := s * s;
      workspace[1] := s * c;
      workspace[2] := (1 - 2 * workspace[0]) / 3;
      workspace[3] := EvaluatePolynomial3(co12, workspace[0]) / (36 * s * c);
      workspace[4] := EvaluatePolynomial4(co13, workspace[0]) / (270 * workspace[0] * c * c);
      x := EvaluatePolynomial5(workspace, eta);
   end
   else begin
      //
      // If eta is large we need to solve Eq 3.2 more directly,
      // begin by getting an initial approximation for x from
      // the last equation on page 155, this is a polynomial in u:
      //
      u := exp(lu);
      workspace[0] := u;
      workspace[1] := alpha;
      workspace[2] := 0;
      workspace[3] := 3 * alpha * (3 * alpha + 1) / 6;
      workspace[4] := 4 * alpha * (4 * alpha + 1) * (4 * alpha + 2) / 24;
      workspace[5] := 5 * alpha * (5 * alpha + 1) * (5 * alpha + 2) * (5 * alpha + 3) / 120;
      x := EvaluatePolynomial6(workspace, u);
      //
      // At this point we may or may not have the right answer, Eq-3.2 has
      // two solutions for x for any given eta, however the mapping in 3.2
      // is 1:1 with the sign of eta and x-sin^2(theta) being the same.
      // So we can check if we have the right root of 3.2, and if not
      // switch x for 1-x.  This transformation is motivated by the fact
      // that the distribution is *almost* symetric so 1-x will be in the right
      // ball park for the solution:
      //
      if ((x - s_2) * eta < 0) then
         x := 1 - x;
   end;
   //
   // The final step is a few Newton-Raphson iterations to
   // clean up our approximation for x, this is pretty cheap
   // in general, and very cheap compared to an incomplete beta
   // evaluation.  The limits set on x come from the observation
   // that the sign of eta and x-sin^2(theta) are the same.
   //
   if (eta < 0) then begin
      lower := 0;
      upper := s_2;
   end
   else begin
      lower := s_2;
      upper := 1;
   end;
   //
   // If our initial approximation is out of bounds then bisect:
   //
   if ((x < lower) or (x > upper)) then
      x := (lower+upper) / 2;
   //
   // And iterate:
   //
   Iter := TTemmeRootFinderIter.Create(-lu,alpha);
   try
     nIter := MaxSeriesIterations;
     NewtonRaphsonIterate(Iter,x,lower,upper,ResNumDigits,nIter);
   finally
     Iter.Free;
   end;

   Result := x;
end;

function TemmeMethod3IBetaInverse(a, b, p, q: double): double;
var
  eta0: double;
  mu,w,w_2,w_3,w_4,w_5,w_6,w_7,w_8,w_9,w_10,d,d_2,d_3,d_4,w1,w1_2,w1_3,w1_4: double;
  u,cross,lower,upper,x: double;
  eta, e1,e2,e3: double;
  nIter: integer;
  Iter: TTemmeRootFinderIter;
begin
   //
   // Begin by getting an initial approximation for the quantity
   // eta from the dominant part of the incomplete beta:
   //
   if (p < q) then
      eta0 := GammaQInvImp(b, p)
   else
      eta0 := GammaPInvImp(b, q);
   eta0 := eta0 / a;
   //
   // Define the variables and powers we'll need later on:
   //
   mu := b / a;
   w := sqrt(1 + mu);
   w_2 := w * w;
   w_3 := w_2 * w;
   w_4 := w_2 * w_2;
   w_5 := w_3 * w_2;
   w_6 := w_3 * w_3;
   w_7 := w_4 * w_3;
   w_8 := w_4 * w_4;
   w_9 := w_5 * w_4;
   w_10 := w_5 * w_5;
   d := eta0 - mu;
   d_2 := d * d;
   d_3 := d_2 * d;
   d_4 := d_2 * d_2;
   w1 := w + 1;
   w1_2 := w1 * w1;
   w1_3 := w1 * w1_2;
   w1_4 := w1_2 * w1_2;
   //
   // Now we need to compute the purturbation error terms that
   // convert eta0 to eta, these are all polynomials of polynomials.
   // Probably these should be re-written to use tabulated data
   // (see examples above), but it's less of a win in this case as we
   // need to calculate the individual powers for the denominator terms
   // anyway, so we might as well use them for the numerator-polynomials
   // as well....
   //
   // Refer to p154-p155 for the details of these expansions:
   //
   e1 := (w + 2) * (w - 1) / (3 * w);
   e1 := e1 + (w_3 + 9 * w_2 + 21 * w + 5) * d / (36 * w_2 * w1);
   e1 := e1 - (w_4 - 13 * w_3 + 69 * w_2 + 167 * w + 46) * d_2 / (1620 * w1_2 * w_3);
   e1 := e1 - (7 * w_5 + 21 * w_4 + 70 * w_3 + 26 * w_2 - 93 * w - 31) * d_3 / (6480 * w1_3 * w_4);
   e1 := e1 - (75 * w_6 + 202 * w_5 + 188 * w_4 - 888 * w_3 - 1345 * w_2 + 118 * w + 138) * d_4 / (272160 * w1_4 * w_5);

   e2 := (28 * w_4 + 131 * w_3 + 402 * w_2 + 581 * w + 208) * (w - 1) / (1620 * w1 * w_3);
   e2 := e2 - (35 * w_6 - 154 * w_5 - 623 * w_4 - 1636 * w_3 - 3983 * w_2 - 3514 * w - 925) * d / (12960 * w1_2 * w_4);
   e2 := e2 - (2132 * w_7 + 7915 * w_6 + 16821 * w_5 + 35066 * w_4 + 87490 * w_3 + 141183 * w_2 + 95993 * w + 21640) * d_2  / (816480 * w_5 * w1_3);
   e2 := e2 - (11053 * w_8 + 53308 * w_7 + 117010 * w_6 + 163924 * w_5 + 116188 * w_4 - 258428 * w_3 - 677042 * w_2 - 481940 * w - 105497) * d_3 / (14696640 * w1_4 * w_6);

   e3 := -((3592 * w_7 + 8375 * w_6 - 1323 * w_5 - 29198 * w_4 - 89578 * w_3 - 154413 * w_2 - 116063 * w - 29632) * (w - 1)) / (816480 * w_5 * w1_2);
   e3 := e3 - (442043 * w_9 + 2054169 * w_8 + 3803094 * w_7 + 3470754 * w_6 + 2141568 * w_5 - 2393568 * w_4 - 19904934 * w_3 - 34714674 * w_2 - 23128299 * w - 5253353) * d / (146966400 * w_6 * w1_3);
   e3 := e3 - (116932 * w_10 + 819281 * w_9 + 2378172 * w_8 + 4341330 * w_7 + 6806004 * w_6 + 10622748 * w_5 + 18739500 * w_4 + 30651894 * w_3 + 30869976 * w_2 + 15431867 * w + 2919016) * d_2 / (146966400 * w1_4 * w_7);
   //
   // Combine eta0 and the error terms to compute eta (Second eqaution p155):
   //
   eta := eta0 + e1 / a + e2 / (a * a) + e3 / (a * a * a);
   //
   // Now we need to solve Eq 4.2 to obtain x.  For any given value of
   // eta there are two solutions to this equation, and since the distribtion
   // may be very skewed, these are not related by x ~ 1-x we used when
   // implementing section 3 above.  However we know that:
   //
   //  cross < x <= 1       ; iff eta < mu
   //          x == cross   ; iff eta == mu
   //     0 <= x < cross    ; iff eta > mu
   //
   // Where cross == 1 / (1 + mu)
   // Many thanks to Prof Temme for clarifying this point.
   //
   // Therefore we'll just jump straight into Newton iterations
   // to solve Eq 4.2 using these bounds, and simple bisection
   // as the first guess, in practice this converges pretty quickly
   // and we only need a few digits correct anyway:
   //
   if (eta <= 0) then
      eta := MinDouble;
   u := eta - mu * ln(eta) + (1 + mu) * ln(1 + mu) - mu;
   cross := 1 / (1 + mu);
   lower := IIF(eta < mu,cross,0);
   upper := IIF(eta < mu,1,cross);
   x := (lower + upper) / 2;

   Iter := TTemmeRootFinderIter.Create(u,mu);
   try
     nIter := MaxSeriesIterations;
     x := NewtonRaphsonIterate(Iter,x,lower,upper,ResNumDigits,nIter);
   finally
     Iter.Free;
   end;
   Result := x;
end;

function IBetaInvImp(a, b, p, q: double; py: PDouble): double;
var
  invert: boolean;
  lower,upper: double;
  l,r,u,x,y: double;
  theta,lambda: double;
  minv,maxv: double;
  ppa,bet,xs,xs2,ps,fs,xg,lx: double;
  ap1,bm1,a_2,a_3,b_2: double;
  digits: integer;
  nIter: integer;
  terms: array[0..4] of double;
  Iter: TIBetaRootsIter;
begin
   // The flag invert is set to true if we swap a for b and p for q,
   // in which case the result has to be subtracted from 1:
   //
   invert := false;
   //
   // Depending upon which approximation method we use, we may end up
   // calculating either x or y initially (where y := 1-x):
   //
   x := 0; // Set to a safe zero to avoid a
   // MSVC 2005 warning C4701: potentially uninitialized local variable 'x' used
   // But code inspection appears to ensure that x IS assigned whatever the code path.

   // For some of the methods we can put tighter bounds
   // on the result than simply [0,1]:
   //
   lower := 0;
   upper := 1;
   //
   // Student's T with b := 0.5 gets handled as a special case, swap
   // around if the arguments are in the "wrong" order:
   //
   if (a = 0.5) then begin
      Swap(a, b);
      Swap(p, q);
      invert := not invert;
   end;
   //
   // Handle trivial cases first:
   //
   if (q = 0) then begin
      if (py <> Nil) then
        py^ := 0;
      Result := 1;
      Exit;
   end
   else if (p = 0) then begin
      if (py <> Nil) then
        py^ := 1;
      Result := 0;
      Exit;
   end
   else if ((a = 1) and (b = 1)) then begin
      if (py <> nil) then
        py^ := 1 - p;
      Result := p;
      Exit;
   end
   else if ((b = 0.5) and (a >= 0.5)) then begin
      //
      // We have a Student's T distribution:
      x := FindIBetaInvFromTDist(a, p, q, @y);
   end
   else if (a + b > 5) then begin
      //
      // When a+b is large then we can use one of Prof Temme's
      // asymptotic expansions, begin by swapping things around
      // so that p < 0.5, we do this to avoid cancellations errors
      // when p is large.
      //
      if (p > 0.5) then begin
         Swap(a, b);
         Swap(p, q);
         invert := not invert;
      end;
      minv := Min(a, b);
      maxv := max(a, b);
      if ((sqrt(minv) > (maxv - minv)) and (minv > 5)) then begin
         //
         // When a and b differ by a small amount
         // the curve is quite symmetrical and we can use an error
         // function to approximate the inverse. This is the cheapest
         // of the three Temme expantions, and the calculated value
         // for x will never be much larger than p, so we don't have
         // to worry about cancellation as long as p is small.
         //
         x := TemmeMethod1IBetaInverse(a, b, p);
         y := 1 - x;
      end
      else begin
         r := a + b;
         theta := ArcSin(sqrt(a / r));
         lambda := minv / r;
         if ((lambda >= 0.2) and (lambda <= 0.8) and (lambda >= 10)) then begin
            //
            // The second error function case is the next cheapest
            // to use, it brakes down when the result is likely to be
            // very small, if a+b is also small, but we can use a
            // cheaper expansion there in any case.  As before x won't
            // be much larger than p, so as long as p is small we should
            // be free of cancellation error.
            //
            ppa := Power(p, 1/a);
            if ((ppa < 0.0025) and (a + b < 200)) then
               x := ppa * Power(a * Beta(a, b), 1/a)
            else
               x := TemmeMethod2IBetaInverse(a, b, p, r, theta);
            y := 1 - x;
         end
         else
         begin
            //
            // If we get here then a and b are very different in magnitude
            // and we need to use the third of Temme's methods which
            // involves inverting the incomplete gamma.  This is much more
            // expensive than the other methods.  We also can only use this
            // method when a > b, which can lead to cancellation errors
            // if we really want y (as we will when x is close to 1), so
            // a different expansion is used in that case.
            //
            if (a < b) then begin
               Swap(a, b);
               Swap(p, q);
               invert := not invert;
            end;
            //
            // Try and compute the easy way first:
            //
            bet := 0;
            if (b < 2) then
               bet := beta(a, b);
            if (bet <> 0) then begin
               y := Power(b * q * bet, 1/b);
               x := 1 - y;
            end
            else
               y := 1;
            if (y > 1e-5) then begin
               x := TemmeMethod3IBetaInverse(a, b, p, q);
               y := 1 - x;
            end
         end
      end
   end
   else if((a < 1) and (b < 1)) then begin
      //
      // Both a and b less than 1,
      // there is a point of inflection at xs:
      //
      xs := (1 - a) / (2 - a - b);
      //
      // Now we need to ensure that we start our iteration from the
      // right side of the inflection point:
      //
      fs := IBeta(a, b, xs) - p;
      if (Abs(fs) / p < const_epsilon * 3) then begin
         // The result is at the point of inflection, best just return it:
         py^ := IIF(invert,xs,1 - xs);
         Result := IIF(invert,1-xs,xs);
         Exit;
      end;
      if (fs < 0) then begin
         Swap(a, b);
         Swap(p, q);
         invert := true;
         xs := 1 - xs;
      end;
      xg := Power(a * p * Beta(a, b), 1/a);
      x := xg / (1 + xg);
      y := 1 / (1 + xg);
      //
      // And finally we know that our result is below the inflection
      // point, so set an upper limit on our search:
      //
      if (x > xs) then
         x := xs;
      upper := xs;
   end
   else if ((a > 1) and (b > 1)) then begin
      //
      // Small a and b, both greater than 1,
      // there is a point of inflection at xs,
      // and it's complement is xs2, we must always
      // start our iteration from the right side of the
      // point of inflection.
      //
      xs := (a - 1) / (a + b - 2);
      xs2 := (b - 1) / (a + b - 2);
      ps := IBeta(a, b, xs) - p;

      if(ps < 0) then begin
         Swap(a, b);
         Swap(p, q);
         Swap(xs, xs2);
         invert := true;
      end;
      //
      // Estimate x and y, using expm1 to get a good estimate
      // for y when it's very small:
      //
      lx := ln(p * a * Beta(a, b)) / a;
      x := exp(lx);

      if x < 0.9 then
        y := 1 - x
      else
        y := -expm1(lx);

      if ((b < a) and (x < 0.2)) then begin
         //
         // Under a limited range of circumstances we can improve
         // our estimate for x, frankly it's clear if this has much effect!
         //
         ap1 := a - 1;
         bm1 := b - 1;
         a_2 := a * a;
         a_3 := a * a_2;
         b_2 := b * b;
         terms[0] := 0;
         terms[1] := 1;
         terms[2] := bm1 / ap1;
         ap1 := ap1 * ap1;
         terms[3] := bm1 * (3 * a * b + 5 * b + a_2 - a - 4) / (2 * (a + 2) * ap1);
         ap1 := ap1 * (a + 1);
         terms[4] := bm1 * (33 * a * b_2 + 31 * b_2 + 8 * a_2 * b_2 - 30 * a * b - 47 * b + 11 * a_2 * b + 6 * a_3 * b + 18 + 4 * a - a_3 + a_2 * a_2 - 10 * a_2)
                    / (3 * (a + 3) * (a + 2) * ap1);
         x := EvaluatePolynomial5(terms, x);
      end;
      //
      // And finally we know that our result is below the inflection
      // point, so set an upper limit on our search:
      //
      if (x > xs) then
         x := xs;
      upper := xs;
   end
   else {if((a <= 1) != (b <= 1))}
   begin
      //
      // If all else fails we get here, only one of a and b
      // is above 1, and a+b is small.  Start by swapping
      // things around so that we have a concave curve with b > a
      // and no points of inflection in [0,1].  As long as we expect
      // x to be small then we can use the simple (and cheap) power
      // term to estimate x, but when we expect x to be large then
      // this greatly underestimates x and leaves us trying to
      // iterate "round the corner" which may take almost forever...
      //
      // We could use Temme's inverse gamma function case in that case,
      // this works really rather well (albeit expensively) even though
      // strictly speaking we're outside it's defined range.
      //
      // However it's expensive to compute, and an alternative approach
      // which models the curve as a distorted quarter circle is much
      // cheaper to compute, and still keeps the number of iterations
      // required down to a reasonable level.  With thanks to Prof Temme
      // for this suggestion.
      //
      if (b < a) then begin
         Swap(a, b);
         Swap(p, q);
         invert := true;
      end;
      if (Power(p, 1/a) < 0.5) then begin
         x := Power(p * a * Beta(a, b), 1 / a);
         if (x = 0) then
            x := MinDouble;
         y := 1 - x;
      end
      else {if(pow(q, 1/b) < 0.1)}
      begin
         // model a distorted quarter circle:
         y := Power(1 - Power(p, b * Beta(a, b)), 1/b);
         if (y = 0) then
            y := MinDouble;
         x := 1 - y;
      end
   end;

   //
   // Now we have a guess for x (and for y) we can set things up for
   // iteration.  If x > 0.5 it pays to swap things round:
   //
   if (x > 0.5) then begin
      Swap(a, b);
      Swap(p, q);
      Swap(x, y);
      invert := not invert;
      l := 1 - upper;
      u := 1 - lower;
      lower := l;
      upper := u;
   end;
   //
   // lower bound for our search:
   //
   // We're not interested in denormalised answers as these tend to
   // these tend to take up lots of iterations, given that we can't get
   // accurate derivatives in this area (they tend to be infinite).
   //
   if (lower = 0) then begin
      if (invert and (py = Nil)) then begin
         //
         // We're not interested in answers smaller than machine epsilon:
         //
         lower := const_epsilon;
         if (x < lower) then
            x := lower;
      end
      else
         lower := MinDouble;
      if (x < lower) then
         x := lower;
   end;
   //
   // Figure out how many digits to iterate towards:
   //
   digits := ResNumDigits div 2;
   if ((x < 1e-50) and ((a < 1) or (b < 1))) then begin
      //
      // If we're in a region where the first derivative is very
      // large, then we have to take care that the root-finder
      // doesn't terminate prematurely.  We'll bump the precision
      // up to avoid this, but we have to take care not to set the
      // precision too high or the last few iterations will just
      // thrash around and convergence may be slow in this case.
      // Try 3/4 of machine epsilon:
      //
      digits := digits * 3;
      digits := digits div 2;
   end;
   //
   // Now iterate, we can use either p or q as the target here
   // depending on which is smaller:
   //
   Iter := TIBetaRootsIter.Create(a,b,IIF(p < q,p,q), IIF(p < q,false,true));
   try
     nIter := MaxSeriesIterations;
     x := HalleyIterate(Iter, x, lower, upper, digits,nIter);
   finally
     Iter.Free;
   end;
   //
   // We don't really want these asserts here, but they are useful for sanity
   // checking that we have the limits right, uncomment if you suspect bugs *only*.
   //
   //BOOST_ASSERT(x != upper);
   //BOOST_ASSERT((x != lower) or (x == boost::math::tools::min_value<T>()) or (x == boost::math::tools::epsilon<T>()));
   //
   // Tidy up, if we "lower" was too high then zero is the best answer we have:
   //
   if (x = lower) then
      x := 0;
   if (py <> Nil) then
      py^ := IIF(invert,x,1 - x);
   Result :=  IIF(invert,1-x,x);
end;


function IBetaInv(a, b, p: double): double;
var
  ry: double;
begin
   Result := IBetaInvImp(a,b,p,1 - p,@ry);
end;

function IBetaC(a, b, x: double): double;
begin
  Result := IBetaImp(a, b, x, True, True, Nil);
end;

function IBeta(a, b, x: double): double;
begin
  Result := IBetaImp(a, b, x, False, True, Nil);
end;

function IBetaDerivativeImp(a, b, x: double): double;
var
  f1: double;
  y: double;
begin
   //
   // Now the corner cases:
   //
   if (x = 0) then begin
     if a > 1 then
       Result := 0
     else begin
       if a = 1 then
         Result := 1 / Beta(a, b)
       else
         raise XLSRWException.Create('IBeta Derivative overlow');
     end;
     Exit;
   end
   else if(x = 1) then begin
     if b > 1 then
       Result := 0
     else begin
       if b = 1 then
         Result := 1 / Beta(a, b)
       else
         raise XLSRWException.Create('IBeta Derivative overlow');
     end;
     Exit;
   end;
   //
   // Now the regular cases:
   //
   f1 := IBetaPowerTerms(a, b, x, 1 - x, Def_Lanczos, true);
   y := (1 - x) * x;

   if( f1 = 0) then begin
     Result := 0;
     Exit;
   end;

   if ((MaxDouble * y < f1)) then
     raise XLSRWException.Create('IBeta Derivative overlow');

   f1 := f1 / y;

   Result := f1;
end;


function GammaPDerivativeImp(a, x: double): double;

var

  f1: double;

begin

   if (x = 0) then
     Result := IIF(a > 1,0,IIF(a = 1,1,MaxDouble))
   else begin
     f1 := RegularisedGammaPrefix(a, x, Def_Lanczos);
     if ((x < 1) and (MaxDouble * x < f1)) then
       Result := MaxDouble
     else
       Result := f1 / x;
   end;
end;

function erf_inv_imp(p, q: double): double;
var
  P_,Q_: array[0..12] of double;
  Y: double;
  R_: double;
  g: double;
  r: double;
  xs: double;
  x: double;
begin
   if (p <= 0.5) then begin
      Y := 0.0891314744949340820313;

       P_[0] :=  -0.000508781949658280665617;
       P_[1] :=  -0.00836874819741736770379;
       P_[2] :=  0.0334806625409744615033;
       P_[3] :=  -0.0126926147662974029034;
       P_[4] :=  -0.0365637971411762664006;
       P_[5] :=  0.0219878681111168899165;
       P_[6] :=  0.00822687874676915743155;
       P_[7] :=  -0.0053877296507124293296;

       Q_[0] :=  1;
       Q_[1] :=  -0.970005043303290640362;
       Q_[2] :=  -1.56574558234175846809;
       Q_[3] :=  1.56221558398423026363;
       Q_[4] :=  0.662328840472002992063;
       Q_[5] :=  -0.71228902341542847553;
       Q_[6] :=  -0.0527396382340099713954;
       Q_[7] :=  0.0795283687341571680018;
       Q_[8] :=  -0.00233393759374190016776;
       Q_[9] :=  0.00088621639045642470750;

      g := p * (p + 10);
      r := EvaluatePolynomial8(P_, p) / EvaluatePolynomial10(Q_, p);
      result := g * Y + g * r;
   end
   else if (q >= 0.25) then begin
      Y := 2.249481201171875;

       P_[0] :=  -0.202433508355938759655;
       P_[1] :=  0.105264680699391713268;
       P_[2] :=  8.37050328343119927838;
       P_[3] :=  17.6447298408374015486;
       P_[4] :=  -18.8510648058714251895;
       P_[5] :=  -44.6382324441786960818;
       P_[6] :=  17.445385985570866523;
       P_[7] :=  21.1294655448340526258;
       P_[8] :=  -3.67192254707729348546;

       Q_[0] :=  1;
       Q_[1] :=  6.24264124854247537712;
       Q_[2] :=  3.9713437953343869095;
       Q_[3] :=  -28.6608180499800029974;
       Q_[4] :=  -20.1432634680485188801;
       Q_[5] :=  48.5609213108739935468;
       Q_[6] :=  10.8268667355460159008;
       Q_[7] :=  -22.6436933413139721736;
       Q_[8] :=  1.7211476576120028272;

      g := sqrt(-2 * ln(q));
      xs := q - 0.25;
      r := EvaluatePolynomial9(P_, xs) / EvaluatePolynomial9(Q_, xs);
      result := g / (Y + r);
   end
   else begin
      x := sqrt(-ln(q));
      if (x < 3) then begin
         // Max error found: 1.089051e-20
         Y := 0.807220458984375;

          P_[0] :=  -0.131102781679951906451;
          P_[1] :=  -0.163794047193317060787;
          P_[2] :=  0.117030156341995252019;
          P_[3] :=  0.387079738972604337464;
          P_[4] :=  0.337785538912035898924;
          P_[5] :=  0.142869534408157156766;
          P_[6] :=  0.0290157910005329060432;
          P_[7] :=  0.00214558995388805277169;
          P_[8] :=  -0.679465575181126350155e-6;
          P_[9] :=  0.285225331782217055858e-7;
          P_[10] :=  -0.681149956853776992068e-9;

          Q_[0] :=  1;
          Q_[1] :=  3.46625407242567245975;
          Q_[2] :=  5.38168345707006855425;
          Q_[3] :=  4.77846592945843778382;
          Q_[4] :=  2.59301921623620271374;
          Q_[5] :=  0.848854343457902036425;
          Q_[6] :=  0.152264338295331783612;
          Q_[7] :=  0.0110592422934648912;

         xs := x - 1.125;
         R_ := EvaluatePolynomial11(P_, xs) / EvaluatePolynomial8(Q_, xs);
         result := Y * x + R_ * x;
      end
      else if (x < 6) then begin
         Y := 0.93995571136474609375;

          P_[0] :=  -0.0350353787183177984712;
          P_[1] :=  -0.00222426529213447927281;
          P_[2] :=  0.0185573306514231072324;
          P_[3] :=  0.00950804701325919603619;
          P_[4] :=  0.00187123492819559223345;
          P_[5] :=  0.000157544617424960554631;
          P_[6] :=  0.460469890584317994083e-5;
          P_[7] :=  -0.230404776911882601748e-9;
          P_[8] :=  0.266339227425782031962e-11;

          Q_[0] :=  1;
          Q_[1] :=  1.3653349817554063097;
          Q_[2] :=  0.762059164553623404043;
          Q_[3] :=  0.220091105764131249824;
          Q_[4] :=  0.0341589143670947727934;
          Q_[5] :=  0.00263861676657015992959;
          Q_[6] :=  0.764675292302794483503e-4;

         xs := x - 3;
         R_ := EvaluatePolynomial9(P_, xs) / EvaluatePolynomial7(Q_, xs);
         result := Y * x + R_ * x;
      end
      else if (x < 18) then begin
         Y := 0.98362827301025390625;

          P_[0] :=  -0.0167431005076633737133;
          P_[1] :=  -0.00112951438745580278863;
          P_[2] :=  0.00105628862152492910091;
          P_[3] :=  0.000209386317487588078668;
          P_[4] :=  0.149624783758342370182e-4;
          P_[5] :=  0.449696789927706453732e-6;
          P_[6] :=  0.462596163522878599135e-8;
          P_[7] :=  -0.281128735628831791805e-13;
          P_[8] :=  0.99055709973310326855e-16;

          Q_[0] :=  1;
          Q_[1] :=  0.591429344886417493481;
          Q_[2] :=  0.138151865749083321638;
          Q_[3] :=  0.0160746087093676504695;
          Q_[4] :=  0.000964011807005165528527;
          Q_[5] :=  0.275335474764726041141e-4;
          Q_[6] :=  0.282243172016108031869e-6;

         xs := x - 6;
         R_ := EvaluatePolynomial9(P_, xs) / EvaluatePolynomial7(Q_, xs);
         result := Y * x + R_ * x;
      end
      else if (x < 44) then begin
         Y := 0.99714565277099609375;

          P_[0] :=  -0.0024978212791898131227;
          P_[1] :=  -0.779190719229053954292e-5;
          P_[2] :=  0.254723037413027451751e-4;
          P_[3] :=  0.162397777342510920873e-5;
          P_[4] :=  0.396341011304801168516e-7;
          P_[5] :=  0.411632831190944208473e-9;
          P_[6] :=  0.145596286718675035587e-11;
          P_[7] :=  -0.116765012397184275695e-17;

          Q_[0] :=  1;
          Q_[1] :=  0.207123112214422517181;
          Q_[2] :=  0.0169410838120975906478;
          Q_[3] :=  0.000690538265622684595676;
          Q_[4] :=  0.145007359818232637924e-4;
          Q_[5] :=  0.144437756628144157666e-6;
          Q_[6] :=  0.509761276599778486139e-9;

         xs := x - 18;
         R_ := EvaluatePolynomial8(P, xs) / EvaluatePolynomial7(Q, xs);
         result := Y * x + R_ * x;
      end
      else begin
         Y := 0.99941349029541015625;

          P_[0] :=  -0.000539042911019078575891;
          P_[1] :=  -0.28398759004727721098e-6;
          P_[2] :=  0.899465114892291446442e-6;
          P_[3] :=  0.229345859265920864296e-7;
          P_[4] :=  0.225561444863500149219e-9;
          P_[5] :=  0.947846627503022684216e-12;
          P_[6] :=  0.135880130108924861008e-14;
          P_[7] :=  -0.348890393399948882918e-21;

          Q_[0] :=  1;
          Q_[1] :=  0.0845746234001899436914;
          Q_[2] :=  0.00282092984726264681981;
          Q_[3] :=  0.468292921940894236786e-4;
          Q_[4] :=  0.399968812193862100054e-6;
          Q_[5] :=  0.161809290887904476097e-8;
          Q_[6] :=  0.231558608310259605225e-1;

         xs := x - 44;
         R_ := EvaluatePolynomial8(P_, xs) / EvaluatePolynomial7(Q_, xs);
         result := Y * x + R_ * x;
      end;
   end;
end;


function erfc_inv(z: double): double;
var
  p, q, s: double;
begin
   if (z > 1) then begin
      q := 2 - z;
      p := 1 - q;
      s := -1;
   end
   else begin
      p := 1 - z;
      q := z;
      s := 1;
   end;
   Result := s * erf_inv_imp(p, q);
end;

function InverseStudentsTTailSeries(df, v: double): double;
var
  w: double;
  np2,np4,np6: double;
  d: array[0..6] of double;
  rn: double;
  div_: double;
  power_: double;
begin
   // Tail series expansion, see section 6 of Shaw's paper.
   // w is calculated using Eq 60:
   w := TGammaDeltaRatioImp(df / 2, 0.5) * Sqrt(df * Pi) * v;
   // define some variables:
   np2 := df + 2;
   np4 := df + 4;
   np6 := df + 6;
   //
   // Calculate the coefficients d(k), these depend only on the
   // number of degrees of freedom df, so at least in theory
   // we could tabulate these for fixed df, see p15 of Shaw:
   //
   d[0] := 1;
   d[1] := -(df + 1) / (2 * np2);
   np2 := np2 *(df + 2);
   d[2] := -df * (df + 1) * (df + 3) / (8 * np2 * np4);
   np2 := np2 * (df + 2);
   d[3] := -df * (df + 1) * (df + 5) * (((3 * df) + 7) * df -2) / (48 * np2 * np4 * np6);
   np2 := np2 *(df + 2);
   np4 := np2 *(df + 4);
   d[4] := -df * (df + 1) * (df + 7) * ( (((((15 * df) + 154) * df + 465) * df + 286) * df - 336) * df + 64 ) / (384 * np2 * np4 * np6 * (df + 8));
   np2 := np2 *(df + 2);
   d[5] := -df * (df + 1) * (df + 3) * (df + 9) * (((((((35 * df + 452) * df + 1573) * df + 600) * df - 2020) * df) + 928) * df -128) / (1280 * np2 * np4 * np6 * (df + 8) * (df + 10));
   np2 := np2 *(df + 2);
   np4 := np2 *(df + 4);
   np6 := np2 *(df + 6);
   d[6] := -df * (df + 1) * (df + 11) * ((((((((((((945 * df) + 31506) * df + 425858) * df + 2980236) * df + 11266745) * df + 20675018) * df + 7747124) * df - 22574632) * df - 8565600) * df + 18108416) * df - 7099392) * df + 884736) / (46080 * np2 * np4 * np6 * (df + 8) * (df + 10) * (df +12));
   //
   // Now bring everthing together to provide the result,
   // this is Eq 62 of Shaw:
   //
   rn := sqrt(df);
   div_ := Power(rn * w, 1 / df);
   power_ := div_ * div_;
   Result := EvaluatePolynomial7(d, power_);
   Result := Result * rn;
   Result := Result / div_;
end;

function InverseStudentsTBodySeries(df, u: double): double;
var
  v: double;
  c: array[0..10] of double;
  in_: double;
begin
   v := TGammaDeltaRatioImp(df / 2, 0.5) * Sqrt(df * Pi * (u - 0.5));
   //
   // Workspace for the polynomial coefficients:
   //
   c[0] := 0;
   c[1] := 0;
   //
   // Figure out what the coefficients are, note these depend
   // only on the degrees of freedom (Eq 57 of Shaw):
   //
   in_ := 1 / df;
   c[2] := 0.16666666666666666667 + 0.16666666666666666667 * in_;
   c[3] := (0.0083333333333333333333 * in_
      + 0.066666666666666666667) * in_
      + 0.058333333333333333333;
   c[4] := ((0.00019841269841269841270 * in_
      + 0.0017857142857142857143) * in_
      + 0.026785714285714285714) * in_
      + 0.025198412698412698413;
   c[5] := (((2.7557319223985890653e-6 * in_
      + 0.00037477954144620811287) * in_
      - 0.0011078042328042328042) * in_
      + 0.010559964726631393298) * in_
      + 0.012039792768959435626;
   c[6] := ((((2.5052108385441718775e-8 * in_
      - 0.000062705427288760622094) * in_
      + 0.00059458674042007375341) * in_
      - 0.0016095979637646304313) * in_
      + 0.0061039211560044893378) * in_
      + 0.0038370059724226390893;
   c[7] := (((((1.6059043836821614599e-10 * in_
      + 0.000015401265401265401265) * in_
      - 0.00016376804137220803887) * in_
      + 0.00069084207973096861986) * in_
      - 0.0012579159844784844785) * in_
      + 0.0010898206731540064873) * in_
      + 0.0032177478835464946576;
   c[8] := ((((((7.6471637318198164759e-13 * in_
      - 3.9851014346715404916e-6) * in_
      + 0.000049255746366361445727) * in_
      - 0.00024947258047043099953) * in_
      + 0.00064513046951456342991) * in_
      - 0.00076245135440323932387) * in_
      + 0.000033530976880017885309) * in_
      + 0.0017438262298340009980;
   c[9] := (((((((2.8114572543455207632e-15 * in_
      + 1.0914179173496789432e-6) * in_
      - 0.000015303004486655377567) * in_
      + 0.000090867107935219902229) * in_
      - 0.00029133414466938067350) * in_
      + 0.00051406605788341121363) * in_
      - 0.00036307660358786885787) * in_
      - 0.00031101086326318780412) * in_
      + 0.00096472747321388644237;
   c[10] := ((((((((8.2206352466243297170e-18 * in_
      - 3.1239569599829868045e-7) * in_
      + 4.8903045291975346210e-6) * in_
      - 0.000033202652391372058698) * in_
      + 0.00012645437628698076975) * in_
      - 0.00028690924218514613987) * in_
      + 0.00035764655430568632777) * in_
      - 0.00010230378073700412687) * in_
      - 0.00036942667800009661203) * in_
      + 0.00054229262813129686486;
   //
   // The result is then a polynomial in v (see Eq 56 of Shaw):
   //
   Result := EvaluateOddPolynomial11(c,v);
end;

procedure HandleZeroDerivative2(Iter: TIterateFuncPair; var last_f0, f0, delta, res, guess, min, max: double);
begin
   if (last_f0 = 0) then begin
      if (res = min) then
         guess := max
      else
         guess := min;
      last_f0 := Iter.Iterate(guess).a;
      delta := guess - res;
   end;
   if (sign(last_f0) * sign(f0) < 0) then begin
      if (delta < 0) then
         delta := (res - min) / 2
      else
         delta := (res - max) / 2;
   end
   else begin
      if (delta < 0) then
         delta := (res - max) / 2
      else
         delta := (res - min) / 2;
   end
end;

procedure HandleZeroDerivative3(Iter: TIterateFuncTriple; var last_f0, f0, delta, res, guess, min, max: double);
begin
   if (last_f0 = 0) then begin
      if (res = min) then
         guess := max
      else
         guess := min;
      last_f0 := Iter.Iterate(guess).a;
      delta := guess - res;
   end;
   if (sign(last_f0) * sign(f0) < 0) then begin
      if (delta < 0) then
         delta := (res - min) / 2
      else
         delta := (res - max) / 2;
   end
   else begin
      if (delta < 0) then
         delta := (res - max) / 2
      else
         delta := (res - min) / 2;
   end
end;

function HalleyIterate(Iter: TIterateFuncTriple; guess, min_, max_: double; digits: integer; var max_iter: integer): double;
var
  f0,f1,f2: double;
  Res3: TFloatTriple;
  factor,delta,last_f0,delta1,delta2: double;
  out_of_bounds_sentry: boolean;
  count: integer;
  denom,num,diff: double;
  convergence: double;
begin
   f0 := 0;
   result := guess;

   factor := ldexp(1.0, 1 - digits);
   delta := Max(10000000 * guess, 10000000);
   last_f0 := 0;
   delta1 := delta;
//   delta2 := delta; unused

   out_of_bounds_sentry := false;

   count  := max_iter;

   repeat
      last_f0 := f0;
      delta2 := delta1;
      delta1 := delta;
      Res3 := Iter.Iterate(result);
      f0 := Res3.a;
      f1 := Res3.b;
      f2 := Res3.c;

      if (0 = f0) then
         break;
      if ((f1 = 0) and (f2 = 0)) then
         HandleZeroDerivative3(Iter, last_f0, f0, delta, result, guess, min_, max_)
      else begin
         if (f2 <> 0) then begin
            denom := 2 * f0;
            num := 2 * f1 - f0 * (f2 / f1);

            if ((Abs(num) < 1) and (Abs(denom) >= Abs(num) * MaxDouble)) then
               delta := f0 / f1
            else
               delta := denom / num;
            if (delta * f1 / f0 < 0) then
               delta := f0 / f1;
         end
         else
            delta := f0 / f1;
      end;
      convergence := Abs(delta / delta2);
      if ((convergence > 0.8) and (convergence < 2)) then begin
         delta := IIF(delta > 0,(result - min_) / 2,(result - max_) / 2);
         if (Abs(delta) > result) then
            delta := sign(delta) * result; // protect against huge jumps!

//         delta2 := delta * 3; Unused
      end;
      guess := result;
      result := result - delta;

      if (result < min_) then
      begin
         diff := IIF((Abs(min_) < 1) and (Abs(result) > 1) and (MaxDouble / Abs(result) < Abs(min_)),1000,result / min_);
         if (Abs(diff) < 1) then
            diff := 1 / diff;
         if(not out_of_bounds_sentry and (diff > 0) and (diff < 3)) then begin
            delta := 0.99 * (guess  - min_);
            result := guess - delta;
            out_of_bounds_sentry := true;
         end
         else begin
            delta := (guess - min_) / 2;
            result := guess - delta;
            if ((result = min_) or (result = max_)) then
               break;
         end;
      end
      else if (result > max_) then begin
         diff := IIF((Abs(max_) < 1) and (Abs(result) > 1) and (MaxDouble / Abs(result) < Abs(max_)),1000,(result / max_));
         if (Abs(diff) < 1) then
            diff := 1 / diff;
         if (not out_of_bounds_sentry and (diff > 0) and (diff < 3)) then begin
            delta := 0.99 * (guess  - max_);
            result := guess - delta;
            out_of_bounds_sentry := true;
         end
         else begin
            delta := (guess - max_) / 2;
            result := guess - delta;
            if ((result = min_) or (result = max_)) then
               break;
         end;
      end;
      if (delta > 0) then
         max_ := guess
      else
         min_ := guess;
      Dec(count)
   until not ((count >= 0) and (Abs(result * factor) < Abs(delta)));

   Dec(max_iter,count);
end;

function InverseStudentsTHill(ndf, u: double): double;
var
   a, b, c, d, q, x, y: double;
begin
   if (ndf > 1e20) then begin
     Result := -erfc_inv(2 * u) * const_root_two;
     Exit;
   end;

   a := 1 / (ndf - 0.5);
   b := 48 / (a * a);
   c := ((20700 * a / b - 98) * a - 16) * a + 96.36;
   d := ((94.5 / (b + c) - 3) / b + 1) * sqrt(a * Pi / 2) * ndf;
   y := Power(d * 2 * u, 2 / ndf);

   if (y > (0.05 + a)) then begin
      //
      // Asymptotic inverse expansion about normal:
      //
      x := -erfc_inv(2 * u) * const_root_two;
      y := x * x;

      if (ndf < 5) then
         c := c + (0.3 * (ndf - 4.5) * (x + 0.6));
      c := c + ((((0.05 * d * x - 5) * x - 7) * x - 2) * x + b);
      y := (((((0.4 * y + 6.3) * y + 36) * y + 94.5) / c - y - 3) / b + 1) * x;
      y := expm1(a * y * y);
   end
   else begin
      y := ((1 / (((ndf + 6) / (ndf * y) - 0.089 * d - 0.822)
              * (ndf + 2) * 3) + 0.5 / (ndf + 4)) * y - 1)
              * (ndf + 1) / (ndf + 2) + 1 / y;
   end;
   q := sqrt(ndf * y);

   Result := -q;
end;

function FindInverseS(p, q: double): double;
const
  a: array[0..3] of double = (3.31125922108741, 11.6616720288968, 4.28342155967104, 0.213623493715853);
  b: array[0..4] of double = (1, 6.61053765625462, 6.40691597760039, 1.27364489782223, 0.3611708101884203e-1);
var
  t,s: double;
begin
   if (p < 0.5) then
      t := sqrt(-2 * ln(p))
   else
      t := sqrt(-2 * ln(q));
   s := t - EvaluatePolynomial4(a, t) / EvaluatePolynomial5(b, t);
   if (p < 0.5) then
      s := -s;
   Result := s;
end;

function DiDonato_SN(a, x: double; N: longword; tolerance: double = 0): double;
var
  i: integer;
  sum, partial: double;
begin
   //
   // Computation of the Incomplete Gamma Function Ratios and their Inverse
   // ARMIDO R. DIDONATO and ALFRED H. MORRIS, JR.
   // ACM Transactions on Mathematical Software, Vol. 12, No. 4,
   // December 1986, Pages 377-393.
   //
   // See equation 34.
   //
   sum := 1;
   if N >= 1 then begin
      partial := x / (a + 1);
      sum := sum + partial;
      for i := 2 to N do begin
         partial := partial * (x / (a + i));
         sum := sum + partial;
         if (partial < tolerance) then
            break;
      end;
   end;
   Result := sum;
end;

function DiDonato_FN(p, a, x: double; N: longword; tolerance: double): double;
var
  u: double;
begin
   u := ln(p) + LGammaImp(a + 1,Def_Lanczos);
   Result := Exp((u + x - ln(DiDonato_SN(a, x, N, tolerance))) / a);
end;

function LGammaImp(z: double; Lanczos: TXLSLanczos; sign: PInteger = Nil): double;
var
  sresult: integer;
  t,zgh: double;
begin
//   static const char* function = "boost::math::lgamma<%1%>(%1%)";

   sresult := 1;
   if (z <= 0) then begin
      // reflection formula:
      if (floor(z) = z) then
         raise XLSRWException.Create('Evaluation of lgamma at a negative integer');

      t := sinpx(z);
      z := -z;
      if (t < 0) then
         t := -t
      else
         sresult := -sresult;
      result := ln(Pi) - LGammaImp(z, Lanczos) - ln(t);
   end
   else if (z < 15) then begin
      result := LGammaSmallImp(z, z - 1, z - 2);
   end
   else if((z >= 3) and (z < 100)) then
      result := ln(GammaImp(z,Lanczos))
   else begin
      zgh := z + Lanczos.g() - 0.5;
      result := ln(zgh) - 1;
      result := result * (z - 0.5);
      result := result + ln(Lanczos.SumExpGScaled(z));
   end;

   if (sign <> Nil) then
      sign^ := sresult;
end;


function FindInverseGamma(a, p, q: double; has_10_digits: PBoolean): double;
var
  g,b,s,t,u,v,w,y,z: double;
  c1,c2,c3,c4,c5,c1_2,c1_3,c1_4,a_2,a_3: double;
  y_2,y_3,y_4: double;
  s_2,s_3,s_4,s_5: double;
  ap1,ap2,ls: double;
  ra: double;
  D_: double;
  lg,lb,zb: double;
begin
   has_10_digits^ := False;
   if a = 1 then begin
      result := -ln(q);
   end
   else if a < 1 then begin
      g := TGamma(a);
      b := q * g;
      if ((b > 0.6) or ((b >= 0.45) and (a >= 0.3))) then begin
         if ((b * q > 1e-8) and (q > 1e-5)) then
            u := Power(p * g * a, 1 / a)
         else
            u := exp((-q / a) - const_euler);
         Result := u / (1 - (u / (a + 1)));
      end
      else if ((a < 0.3) and (b >= 0.35)) then begin
         t := exp(const_euler - b);
         u := t * exp(t);
         result := t * exp(u);
      end
      else if ((b > 0.15) or (a >= 0.3)) then begin
         y := -ln(b);
         u := y - (1 - a) * ln(y);
         result := y - (1 - a) * ln(u) - ln(1 + (1 - a) / (1 + u));
      end
      else if (b > 0.1) then begin
         // DiDonato & Morris Eq 24:
         y := -ln(b);
         u := y - (1 - a) * ln(y);
         result := y - (1 - a) * ln(u) - ln((u * u + 2 * (3 - a) * u + (2 - a) * (3 - a)) / (u * u + (5 - a) * u + 2));
      end
      else begin
         // DiDonato & Morris Eq 25:
         y := -ln(b);
         c1 := (a - 1) * ln(y);
         c1_2 := c1 * c1;
         c1_3 := c1_2 * c1;
         c1_4 := c1_2 * c1_2;
         a_2 := a * a;
         a_3 := a_2 * a;

         c2 := (a - 1) * (1 + c1);
         c3 := (a - 1) * (-(c1_2 / 2) + (a - 2) * c1 + (3 * a - 5) / 2);
         c4 := (a - 1) * ((c1_3 / 3) - (3 * a - 5) * c1_2 / 2 + (a_2 - 6 * a + 7) * c1 + (11 * a_2 - 46 * a + 47) / 6);
         c5 := (a - 1) * (-(c1_4 / 4)
                           + (11 * a - 17) * c1_3 / 6
                           + (-3 * a_2 + 13 * a -13) * c1_2
                           + (2 * a_3 - 25 * a_2 + 72 * a - 61) * c1 / 2
                           + (25 * a_3 - 195 * a_2 + 477 * a - 379) / 12);

         y_2 := y * y;
         y_3 := y_2 * y;
         y_4 := y_2 * y_2;
         Result := y + c1 + (c2 / y) + (c3 / y_2) + (c4 / y_3) + (c5 / y_4);
         if (b < 1e-28) then
            has_10_digits^ := true;
      end;
   end
   else begin
      s := FindInverseS(p, q);

      s_2 := s * s;
      s_3 := s_2 * s;
      s_4 := s_2 * s_2;
      s_5 := s_4 * s;
      ra := sqrt(a);

      w := a + s * ra + (s * s -1) / 3;
      w := w + (s_3 - 7 * s) / (36 * ra);
      w := w - (3 * s_4 + 7 * s_2 - 16) / (810 * a);
      w := w + (9 * s_5 + 256 * s_3 - 433 * s) / (38880 * a * ra);

      if ((a >= 500) and (abs(1 - w / a) < 1e-6)) then begin
         Result := w;
         has_10_digits^ := true;
      end
      else if (p > 0.5) then begin
         if (w < 3 * a) then
            result := w
         else begin
            D_ := Max(2, a * (a - 1));
            lg := LGammaImp(a,Def_Lanczos);
            lb := ln(q) + lg;
            if (lb < -D_ * 2.3) then begin
               // DiDonato and Morris Eq 25:
               y := -lb;
               c1 := (a - 1) * ln(y);
               c1_2 := c1 * c1;
               c1_3 := c1_2 * c1;
               c1_4 := c1_2 * c1_2;
               a_2 := a * a;
               a_3 := a_2 * a;

               c2 := (a - 1) * (1 + c1);
               c3 := (a - 1) * (-(c1_2 / 2) + (a - 2) * c1 + (3 * a - 5) / 2);
               c4 := (a - 1) * ((c1_3 / 3) - (3 * a - 5) * c1_2 / 2 + (a_2 - 6 * a + 7) * c1 + (11 * a_2 - 46 * a + 47) / 6);
               c5 := (a - 1) * (-(c1_4 / 4)
                                 + (11 * a - 17) * c1_3 / 6
                                 + (-3 * a_2 + 13 * a -13) * c1_2
                                 + (2 * a_3 - 25 * a_2 + 72 * a - 61) * c1 / 2
                                 + (25 * a_3 - 195 * a_2 + 477 * a - 379) / 12);

               y_2 := y * y;
               y_3 := y_2 * y;
               y_4 := y_2 * y_2;
               Result := y + c1 + (c2 / y) + (c3 / y_2) + (c4 / y_3) + (c5 / y_4);
            end
            else
            begin
               u := -lb + (a - 1) * ln(w) - ln(1 + (1 - a) / (1 + w));
               result := -lb + (a - 1) * ln(u) - ln(1 + (1 - a) / (1 + u));
            end
         end
      end
      else
      begin
         z := w;
         ap1 := a + 1;
         ap2 := a + 2;
         if (w < 0.15 * ap1) then begin
          // DiDonato and Morris Eq 35:
          v := ln(p) + LGammaImp(ap1,Def_Lanczos);
          z := exp((v + w) / a);
          s := log1p(z / ap1 * (1 + z / ap2));
          z := exp((v + z - s) / a);
          s := log1p(z / ap1 * (1 + z / ap2));
          z := exp((v + z - s) / a);
          s := log1p(z / ap1 * (1 + z / ap2 * (1 + z / (a + 3))));
          z := exp((v + z - s) / a);
         end;

         if ((z <= 0.01 * ap1) or (z > 0.7 * ap1)) then begin
          result := z;
          if (z <= 0.002 * ap1) then
             has_10_digits^ := true;
         end
         else begin
          // DiDonato and Morris Eq 36:
          ls := ln(didonato_SN(a, z, 100, 1e-4));
          v := ln(p) + LGammaImp(ap1, Def_Lanczos);
          z := exp((v + z - ls) / a);
          result := z * (1 - (a * ln(z) - z - v + ls) / (a - z));
         end;
      end;
   end;
end;

function GammaPInvImp(a, p: double): double;
var
  guess: double;
  lower: double;
  digits: integer;
  Iter: TGammaPInverseIter;
  Has10Digits: boolean;
  nIter: integer;
begin
//   static const char* function = "boost::math::gamma_p_inv<%1%>(%1%, %1%)";

   if p = 1 then begin
     Result := MaxDouble;
     Exit;
   end;
   if p = 0 then begin
     Result := 0;
     Exit;
   end;
   guess := FindInverseGamma(a, p, 1 - p,@Has10Digits);
   lower := MinDouble;
   if guess <= lower then
      guess := MinDouble;

   //
   // Work out how many digits to converge to, normally this is
   // 2/3 of the digits in T, but if the first derivative is very
   // large convergence is slow, so we'll bump it up to full
   // precision to prevent premature termination of the root-finding routine.
   //
   digits := (ResNumDigits * 2) div 3;
   if ((a < 0.125) and (Abs(GammaPDerivativeImp(a, guess)) > 1 / sqrt(const_epsilon))) then
      digits := ResNumDigits - 2;
   //
   // Go ahead and iterate:
   //
   Iter := TGammaPInverseIter.Create(a,p,False);
   try
     nIter := MaxSeriesIterations;
     guess := HalleyIterate(Iter,guess, lower,MaxDouble, digits, nIter);
   finally
     Iter.Free;
   end;
   if (guess = lower) then
      raise XLSRWException.Create('Expected result known to be non-zero, but is smaller than the smallest available number.');
   Result := guess;
end;

function BetaCDF(A, B, X: double; Inverse: boolean = False): double;
begin
  if (X = 0) or (X = 1) then
    Result := X
  else
    Result := IBetaImp(A,B,X,Inverse,True,Nil);
end;

//function BetaInv(a, b, p: double): double;
//begin
//  if p = 0 then
//    Result := 0
//  else if p = 1 then
//    Result := 1
//  else
//    Result := IBetaInv(a, b, p);
//end;

function BetaPDF(a, b, x: double): double;
begin
  Result := IBetaDerivativeImp(a,b,x);
end;

function BetaInvCDF(a, b, x: double): double;
begin
  Result := IBetaImp(A,B,X,True,True,Nil);
end;

function HypergeometricPdfFactorialImp(x, r, n, N_: longword): double;
var
  i: integer;
  j: integer;
  num: array[0..2] of double;
  denom: array[0..4] of double;
begin
   result := QuickFactorial(n);
      num[0] := QuickFactorial(r);
      num[1] := QuickFactorial(N_ - n);
      num[2] := QuickFactorial(N_ - r);

      denom[0] := QuickFactorial(N_);
      denom[1] := QuickFactorial(x);
      denom[2] := QuickFactorial(n - x);
      denom[3] := QuickFactorial(r - x);
      denom[4] := QuickFactorial(N_ - n - r + x);


   i := 0;
   j := 0;
   while ((i < 3) or (j < 5)) do begin
      while ((j < 5) and ((result >= 1) or (i >= 3))) do begin
         result := result / denom[j];
         Inc(j);
      end;
      while ((i < 3) and ((result <= 1) or (j >= 5))) do begin
         result := result * num[i];
         Inc(i);
      end
   end
end;

function HypergeometricPdfPrimeImp(x, r, n, N_: longword): double;
begin
  Result := 0;
end;

function HypergeometricPdfLanczosImp(x, r, n, N_: longword): double;
begin
  Result := 0;
end;

function HypergeometricDistCDF(x, r, n, N_: longword): double;
begin
  Result := HypergeometricDistCDFImp(x,r,n,N_,False);
  if(Result > 1) then
    Result := 1
  else if (result < 0) then
    Result := 0;
end;

function HypergeometricDistPDF(x, r, n, N_: longword): double;
begin
  Result := HypergeometricDistPDFImp(x, r, n, N_);
end;

function HypergeometricDistCDFImp(x, r, n, N_: longword; invert: boolean): double;
var
  mode: double;
  diff: double;
  lower_limit: longword;
  upper_limit: longword;
begin
  result := 0;
  mode := floor(r + 1 * n + 1 / (N_ + 2));
  if (x < mode) then begin
     result := HypergeometricDistPDF(x, r, n, N_);
     diff := result;
     lower_limit := Max(0,(n + r) - N_);
     while (diff > IIF(invert,1,result) * const_epsilon) do begin
        diff := x * ((N_ + x) - n - r) * diff / ((1 + n - x) * (1 + r - x));
        result := result + diff;
        if (x = lower_limit) then
          break;
        Dec(x);
     end
  end
  else
  begin
     invert := not invert;
     upper_limit := Min(r, n);
     if (x <> upper_limit) then begin
        Inc(x);
        result := HypergeometricDistPDF(x, r, n, N_);
        diff := result;
        while((x <= upper_limit) and (diff > IIF(invert,1,result) * const_epsilon)) do begin
           diff := (n - x) * (r - x) * diff / ((x + 1) * ((N_ + x + 1) - n - r));
           result := result + diff;
           Inc(x);
        end;
     end;
  end;
  if (invert) then
     result := 1 - result;
end;

function HypergeometricDistPDFImp(x, r, n, N_: longword): double;
begin
   if (N_ <= MAX_QUICKFACTORIAL) then
      //
      // If N is small enough then we can evaluate the PDF via the factorials
      // directly: table lookup of the factorials gives the best performance
      // of the methods available:
      //
      result := HypergeometricPdfFactorialImp(x, r, n, N_)
//   else if (N_ <= max_prime - 1 ) then Prime list not implemented
//      //
//      // If N is no larger than the largest prime number in our lookup table
//      // (104729) then we can use prime factorisation to evaluate the PDF,
//      // this is slow but accurate:
//      //
//      result := HypergeometricPdfPrimeImp(x, r, n, N_)
   else
      //
      // Catch all case - use the lanczos approximation - where available -
      // to evaluate the ratio of factorials.  This is reasonably fast
      // (almost as quick as using logarithmic evaluation in terms of lgamma)
      // but only a few digits better in accuracy than using lgamma:
      //
      result := HypergeometricPdfLanczosImp(x, r, n, N_{, evaluation_type()});

   if (result > 1) then
      result := 1
   else if (result < 0) then
      result := 0;
end;

function BinomialCDF(k: longword; n: longword; p: double): double;
begin
  if (k = n) then
    Result := 1
  else if (p = 0) then
     Result := 1
  else if (p = 1) then
    Result := 0
  else
    Result := IBetaImp(k + 1, n - k, p,True,True,Nil);
end;

function BinomialCDF2(k: longword; n: longword; p: double): double;
begin
  if (k = n) then
    Result := 1
  else if (p = 0) then
     Result := 1
  else if (p = 1) then
    Result := 0
  else
    Result := IBetaImp(k + 1, n - k, p,False,True,Nil);
end;

function BinomialPDF(k: longword; n: longword; p: double): double;
begin
  if (p = 0) then
    Result := IIF(k = 0,1,0)
  else if (p = 1) then
    Result := IIF(k = n,1,0)
  else if (n = 0) then
    Result := 1
  else if (k = 0) then
    Result := Power(1 - p, n)
  else if (k = n) then
    Result := Power(p, k)
  else
    Result := IBetaDerivativeImp(k + 1,n - k + 1,p) / (n + 1);
end;

function BinomialQuantile(n: integer; p,p2: double): double;
begin
  Result := BinomialQuantileImp(n,p2,p,1 - p);
end;

function ChiSquaredDistributionCDF(DegreesOfFreedom,ChiSquare: double): double;
begin
  // Not sure if this is ok, but it makes excel happy.
  if ChiSquare = 0 then
    Result := 0
  else
    Result := Gamma_P(DegreesOfFreedom / 2, ChiSquare / 2);
end;

function ChiSquaredDistributionPDF(DegreesOfFreedom,ChiSquare: double): double;
begin
  if (ChiSquare = 0) then begin
    // Handle special cases:
    if (DegreesOfFreedom < 2) then
      Result := MaxDouble
    else if (DegreesOfFreedom = 2) then
      Result := 0.5
    else
      Result := 0;
  end
  else
    Result := Gamma_Q(DegreesOfFreedom / 2, ChiSquare / 2);
end;

function ChiSquaredDistributionInvPDF(df,x: double): double;
var
  Scale: double;
begin
   Scale := 1;
   if x = 0 then
     Result := 0
   else begin
     Result := df * scale / 2 / x;
     Result := GammaPDerivativeImp(df / 2, Result) * df * scale / 2;
     if Result <> 0 then
        Result := Result / (x * x);
   end;
end;

function GaussInvPDF(x,mean: double): double;
var
  scale: double;
begin
  scale := 1;
  Result := sqrt(scale / ((Pi / 2) * x * x * x)) * exp(-scale * (x - mean) * (x - mean) / (2 * x * mean * mean));
end;

function ExponentialDistCDF(x,lambda: double): double;
begin
  Result := -expm1(-x * lambda);
end;

function ExponentialDistPDF(x,lambda: double): double;
begin
  Result := lambda * exp(-lambda * x);
end;

function GammaCDF(x,shape,scale: double): double;
begin
  Result := 1 - Gamma_Q(shape, x / scale);
end;

function GammaPDF(x,shape,scale: double): double;
begin
  Result := GammaPDerivativeImp(shape, x / scale) / scale;
end;

function GammaDistributionQuantile(a,b,p: double): double;
begin
  Result := GammaPInvImp(a,p) * b;
end;

function NormalCDF(x, mean, stdev: double): double;
var
  diff: double;
begin
  diff := (x - mean) / (stdev * const_root_two);
  Result := erfc(-diff) / 2;
end;

function NormalPDF(x, mean, stdev: double): double;
var
  exponent: double;
begin
  exponent := x - mean;
  exponent := exponent * -exponent;
  exponent := exponent / (2 * stdev * stdev);

  result := Exp(exponent);
  result := result / (stdev * sqrt(2 * Pi));
end;

function LogNormalCDF(x, mean, stdev: double): double;
begin
  Result := 0;
end;

function NormalQuantile(p, mean, stdev: double): double;
begin
  Result := erfc_inv(2 * p);
  Result := -result;
  Result := Result * stdev * const_root_two;
  Result := Result + mean;
end;

function LogNormalPDF(x, mean, stdev: double): double;
var
  exponent: double;
begin
   exponent := ln(x) - mean;
   exponent := exponent * -exponent;
   exponent := exponent / (2 * stdev * stdev);

   result := Exp(exponent);
   result := result / (stdev * sqrt(2 * Pi) * x);
end;

function LogNormalCDFInv(x, mean, stdev: double): double;
begin
  Result := 0;
end;

function LogNormalQuantile(p, mean, stdev: double): double;
begin
 if p = 0 then
   Result := 0
 else if p = 1 then
   raise XLSRWException.Create('Overflow')
 else
   Result := Exp(NormalQuantile(p,mean,stdev));
end;

function NegBinomialCDF(k, r, p: double): double;
begin
  Result := IBeta(r,k + 1, p);
end;

function NegBinomialPDF(k, r, p: double): double;
begin
  Result := (p /(r + k)) * IBetaDerivativeImp(r, k +1, p);
end;

function PoissonCDF(k, mean: double): double;
begin
  if mean = 0 then
    Result := 0
  else if k = 0 then
    Result := Exp(-mean)
  else
    Result := Gamma_Q(k + 1, mean);
end;

function PoissonPDF(k, mean: double): double;
begin
  if mean = 0 then
    Result := 0
  else if k = 0 then
    Result := Exp(-mean)
  else
    Result := GammaPDerivativeImp(k + 1, mean);
end;

function StudentTDistCDF(t ,DegreesOfFreedom: double): double;
var
  t2: double;
  z: double;
  probability: double;
begin
  t2 := t * t;
  if DegreesOfFreedom > 2 * t2 then begin
    z := t2 / (DegreesOfFreedom + t2);
    probability := IBetaC(0.5, DegreesOfFreedom / 2, z) / 2;
  end
  else begin
     z := DegreesOfFreedom / (DegreesOfFreedom + t2);
     probability := IBeta(DegreesOfFreedom / 2, 0.5, z) / 2;
  end;
  Result := IIF(t > 0,1 - probability,probability);
end;

function StudentTDistPDF(t ,DegreesOfFreedom: double): double;
var
  basem1: double;
begin
  basem1 := t * t / DegreesOfFreedom;
  if basem1 < 0.125 then
    Result := Exp(-log1p(basem1) * (1 + DegreesOfFreedom) / 2)
  else
    Result := Power(1 / (1 + basem1), (DegreesOfFreedom + 1) / 2);
  Result := Result / (Sqrt(DegreesOfFreedom) * Beta(DegreesOfFreedom / 2, 0.5));
end;

function WeibullCDF(x, shape,scale: double): double;
begin
  Result := -expm1(-Power(x / scale, shape));
end;

function WeibullPDF(x, shape,scale: double): double;
begin
  if x = 0 then
    Result := 0
  else begin
    Result := Exp(-Power(x / scale, shape));
    Result := Result * (Power(x / scale, shape) * shape / x);
  end;
end;

function InverseStudentsT(df, u, v: double; pexact: PBoolean = Nil): double;
var
  _v: integer;
  invert: boolean;
//  tolerance: double;
  crossover: double;
label
  calculate_real;
begin
   //
   // df = number of degrees of freedom.
   // u = probablity.
   // v = 1 - u.
   // l = lanczos type to use.
   //
   invert := false;
   if pexact <> Nil then
      pexact^ := false;
   if (u > v) then begin
      // function is symmetric, invert it:
      Swap(u, v);
      invert := true;
   end;
   if ((floor(df) = df) and (df < 20)) then begin
      //
      // we have integer degrees of freedom, try for the special
      // cases first:
      //
//      tolerance := ldexp(1.0, (2 * ResNumDigits) div 3);
//      nor required since case 5 & 6 are removed (more code to translate...)

      _v := itrunc(df);
      case _v of
         1: begin
            //
            // df := 1 is the same as the Cauchy distribution, see
            // Shaw Eq 35:
            //
            if (u = 0.5) then
               result := 0
            else
               Result := -cos(Pi * u) / sin(Pi * u);
            if (pexact <> Nil) then
               pexact^ := true;
         end;
         2: begin
            //
            // df := 2 has an exact result, see Shaw Eq 36:
            //
            result := (2 * u - 1) / sqrt(2 * u * v);
            if (pexact <> Nil) then
               pexact^ := true;
         end;
      else goto calculate_real;
      end
   end
   else begin
calculate_real:
      if (df < 3) then begin
         //
         // Use a roughly linear scheme to choose between Shaw's
         // tail series and body series:
         //
         crossover := 0.2742 - df * 0.0242143;
         if (u > crossover) then
            result := InverseStudentsTBodySeries(df, u)
         else
            result := InverseStudentsTTailSeries(df, u);
      end
      else begin
         //
         // Use Hill's method except in the exteme tails
         // where we use Shaw's tail series.
         // The crossover point is roughly exponential in -df:
         //
         crossover := ldexp(1.0, Round(df / -0.654));
         if (u > crossover) then
            result := InverseStudentsTHill(df, u)
         else
            result := InverseStudentsTTailSeries(df, u);
      end;
   end;
   Result := IIF(invert,-result,result);
end;

function ChiSquaredQuantile(DegreesOfFreedom, p: double): double;
begin
  Result := 2 * GammaPInvImp(DegreesOfFreedom / 2, p);
end;

function FisherFDistCDF(x, df1, df2: double): double;
var
  v1x: double;
begin
   v1x := df1 * x;
   //
   // There are two equivalent formulas used here, the aim is
   // to prevent the final argument to the incomplete beta
   // from being too close to 1: for some values of df1 and df2
   // the rate of change can be arbitrarily large in this area,
   // whilst the value we're passing will have lost information
   // content as a result of being 0.999999something.  Better
   // to switch things around so we're passing 1-z instead.
   //

   if v1x > df2 then
     Result := IBetaC(df2 / 2, df1 / 2, df2 / (df2 + v1x))
   else
     Result := IBeta(df1 / 2, df2 / 2, v1x / (df2 + v1x));
end;

function FisherFDistPDF(x, df1, df2: double): double;
var
  v1x: double;
begin
   if (x = 0) then begin
      // special cases:
      if (df1 < 2) then
        raise XLSRWException.Create('FisherFDistPDF Overflow')
      else if (df1 = 2) then
         Result := 1
      else
         Result := 0;
      Exit;
   end;

   //
   // You reach this formula by direct differentiation of the
   // cdf expressed in terms of the incomplete beta.
   //
   // There are two versions so we don't pass a value of z
   // that is very close to 1 to ibeta_derivative: for some values
   // of df1 and df2, all the change takes place in this area.
   //
   v1x := df1 * x;
   if (v1x > df2) then begin
      Result := (df2 * df1) / ((df2 + v1x) * (df2 + v1x));
      Result := Result * IbetaDerivativeImp(df2 / 2, df1 / 2, df2 / (df2 + v1x));
   end
   else
   begin
      Result := df2 + df1 * x;
      Result := (result * df1 - x * df1 * df1) / (result * result);
      Result := Result * IbetaDerivativeImp(df1 / 2, df2 / 2, v1x / (df2 + v1x));
   end;
end;

function FisherFDistQuantile(p, df1, df2: double): double;
var
  x,y: double;
begin
  x := IBetaInvImp(df1 / 2, df2 / 2, p, 1 - p, @y);

 Result := df2 * x / (df1 * y);
end;

{ TXLSLanczos13m53 }

function TXLSLanczos13m53.g: double;
begin
  Result := 6.024680040776729583740234375;
end;

function TXLSLanczos13m53.Sum(z: double): double;
const
  Num: array[0..12] of double =  (
  23531376880.41075968857200767445163675473,
  42919803642.64909876895789904700198885093,
  35711959237.35566804944018545154716670596,
  17921034426.03720969991975575445893111267,
  6039542586.35202800506429164430729792107,
  1439720407.311721673663223072794912393972,
  248874557.8620541565114603864132294232163,
  31426415.58540019438061423162831820536287,
  2876370.628935372441225409051620849613599,
  186056.2653952234950402949897160456992822,
  8071.672002365816210638002902272250613822,
  210.8242777515793458725097339207133627117,
  2.506628274631000270164908177133837338626);

  Denom: array[0..12] of longword = (
  0,
  39916800,
  120543840,
  150917976,
  105258076,
  45995730,
  13339535,
  2637558,
  357423,
  32670,
  1925,
  66,
  1);
begin
  Result := EvaluateRational(Num,Denom,Length(Num),z);
end;

function TXLSLanczos13m53.SumExpGScaled(z: double): double;
const
  Num: array[0..12] of double = (
  56906521.91347156388090791033559122686859,
  103794043.1163445451906271053616070238554,
  86363131.28813859145546927288977868422342,
  43338889.32467613834773723740590533316085,
  14605578.08768506808414169982791359218571,
  3481712.15498064590882071018964774556468,
  601859.6171681098786670226533699352302507,
  75999.29304014542649875303443598909137092,
  6955.999602515376140356310115515198987526,
  449.9445569063168119446858607650988409623,
  19.51992788247617482847860966235652136208,
  0.5098416655656676188125178644804694509993,
  0.00606184234624890652578375396455593688322);

  Denom: array[0..12] of longword = (
  0,
  39916800,
  120543840,
  150917976,
105258076,
  45995730,
  13339535,
  2637558,
  357423,
  32670,
  1925,
  66,
  1);

begin
  Result := EvaluateRational(Num,Denom,Length(Num),z);
end;

{ TIBetaSeriesT }

constructor TIBetaSeriesIter.Create(a, b, x, mult: double);
begin
  FRes := mult;
  Fx := x;
  Fapn := a;
  Fpoch := 1 - b;
  Fn := 1;
end;

function TIBetaSeriesIter.Iterate: double;
var
  r: double;
begin
  r := FRes / Fapn;
  Fapn := Fapn + 1;
  FRes := FRes * (Fpoch * Fx / Fn);
  Inc(Fn);
  Fpoch := Fpoch + 1;
  Result := r;
end;

{ TLowerIncompleteGammaSeries }

constructor TLowerIncompleteGammaSeriesIter.Create(a,z: double);
begin
  Fa := a;
  Fz := z;
  FRes := 1;
end;

function TLowerIncompleteGammaSeriesIter.Iterate: double;
begin
  Result := FRes;
  Fa := Fa + 1;
  FRes := FRes * (Fz / Fa);
end;

{ TSmallGamma2Series }

constructor TSmallGamma2SeriesIter.Create(a,x: double);
begin
  FRes := -x;
  Fx := -x;
  Fapn := a + 1;
  Fn := 1;
end;

function TSmallGamma2SeriesIter.Iterate: double;
var
  r: double;
begin
  Inc(Fn);
  r := FRes / (Fapn);
  FRes := FRes * Fx;
  FRes := FRes / (Fn + 1);
  Fapn := Fapn + 1;
  Result := r;
end;

{ TErfAsymptSeries }

constructor TErfAsymptSeriesIter.Create(z: double);
begin
  Fxx := 2 * -z * z;
  Ftk := 1;
end;

function TErfAsymptSeriesIter.Iterate: double;
var
  r: double;
begin
  r := FRes;
  FRes := FRes * (Ftk / Fxx);
  Inc(Ftk,2);
  if (Abs(r) < Abs(FRes)) then
    FRes := 0;
  Result := r;
end;

{ IBetaFraction2 }

constructor TIBetaFraction2Iter.Create(a, b, x: double);
begin
  Fa := a;
  Fb := b;
  Fx := x;
  Fm := 0;
end;

function TIBetaFraction2Iter.Iterate: TFloatPair;
var
  aN: double;
  bN: double;
  denom: double;
begin
	  aN := (Fa + Fm - 1) * (Fa + Fb + Fm - 1) * Fm * (Fb - Fm) * Fx * Fx;
	  denom := (Fa + 2 * Fm - 1);
	  aN := aN / (denom * denom);

	  bN := Fm;
	  bN := bN + (Fm * (Fb - Fm) * Fx) / (Fa + 2*Fm - 1);
	  bN := bN + ((Fa + Fm) * (Fa - (Fa + Fb) * Fx + 1 + Fm *(2 - Fx))) / (Fa + 2*Fm + 1);

	  Inc(Fm);

    Result.a := aN;
    Result.b := bN;
end;

{ TUpperIncompleteGammaFractIter }

constructor TUpperIncompleteGammaFractIter.Create(a, z: double);
begin
  Fa := a;
  Fz := z - a + 1;
  Fk := 0;
end;

function TUpperIncompleteGammaFractIter.Iterate: TFloatPair;
begin
  Inc(Fk);
  Fz := Fz + 2;
  Result.a := Fk * (Fa - Fk);
  Result.b := Fz;
end;

{ TGammaPInverseIter }

constructor TGammaPInverseIter.Create(Aa, Ap: double; AInv: boolean);
begin
  Fa := Aa;
  Fp := Ap;
  FInv := AInv;
  if Fp > 0.9 then begin
     Fp := 1 - Fp;
     FInv := not FInv;
  end;
end;

function TGammaPInverseIter.Iterate(x: double): TFloatTriple;
var
  f,f1,f2,div_,ft: double;
begin
  ft := 0;
  f := GammaIncompleteImp(Fa,x, true, FInv, @ft);
  f1 := ft;
  div_ := (Fa - x - 1) / x;
  f2 := f1;
  if ((Abs(div_) > 1) and (MaxDouble / Abs(div_) < f2)) then
     f2 := -MaxDouble / 2
  else
     f2 := f2 * div_;

  if (FInv) then begin
     f1 := -f1;
     f2 := -f2;
  end;

  Result.a := f - Fp;
  Result.b := f1;
  Result.c := f2;
end;

{ TTemmeRootFinderIter }

constructor TTemmeRootFinderIter.Create(At, Aa: double);
begin
  Ft := At;
  Fa := Aa;
end;

function TTemmeRootFinderIter.Iterate(x: double): TFloatPair;
var
  y,f,f1: double;
  big: double;
begin
  y := 1 - x;
  if (y = 0) then begin
     big := MaxDouble / 4;
     Result.a := -big;
     Result.b := -big;
     Exit;
  end;
  if (x = 0) then begin
     big := MaxDouble / 4;
     Result.a := -big;
     Result.b := big;
     Exit;
  end;
  f := ln(x) + Fa * ln(y) + Ft;
  f1 := (1 / x) - (Fa / (y));
  Result.a := -f;
  Result.b := f1;
end;

{ TIBetaRootsIter }

constructor TIBetaRootsIter.Create(Aa, Ab, At: double; AInverse: boolean);
begin
  Fa := Aa;
  Fb := Ab;
  FTarget := At;
  FInvert := AInverse;
end;

function TIBetaRootsIter.Iterate(x: double): TFloatTriple;
var
  f,f1,f2,y: double;
begin
  y := 1 - x;
  f := IBetaImp(Fa, Fb, x, Finvert, true, @f1) - Ftarget;
  if (Finvert) then
   f1 := -f1;
  if (y = 0) then
   y := MinDouble * 64;
  if (x = 0) then
   x := MinDouble * 64;

  f2 := f1 * (-y * Fa + (Fb - 2) * x + 1);
  if (Abs(f2) < y * x * MaxDouble) then
   f2 := f2 / (y * x);
  if (Finvert) then
   f2 := -f2;

  // make sure we don't have a zero derivative:
  if (f1 = 0) then
   f1 := IIF(Finvert,-1,1) * MinDouble * 64;

  Result.a := f;
  Result.b := f1;
  Result.c := f2;
end;

{ DistributionQuantileIter }

constructor DistributionQuantileIter.Create(n: longword; success_fraction, p, q: double);
begin
  FComp := IIF(p < q,False,True);
  FTarget := IIF(p < q,p,q);
  Fn := n;
  Fsuccess_fraction := success_fraction;
end;

function DistributionQuantileIter.Iterate(x: double): double;
begin
  if FComp then
    Result := FTarget - BinomialCDF2(Trunc(x),Trunc(Fn),Fsuccess_fraction)
  else
    Result := BinomialCDF(Trunc(x),Trunc(Fn),Fsuccess_fraction) - FTarget;
end;

initialization
  Def_Lanczos := TXLSLanczos13m53.Create;

finalization
  Def_Lanczos.Free;

end.
