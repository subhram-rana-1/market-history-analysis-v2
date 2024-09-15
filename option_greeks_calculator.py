import math
from scipy.stats import norm

# S0 = underlying price
# X = strike price
# σ = volatility
# r = continuously compounded risk-free interest rate
# q = continuously compounded dividend yield
# t = time to expiration
# For,
# σ = Volatility = India VIX has been taken.
# r = 10% (As per NSE Website, it is fixed.)
# q = 0.00% (Assumed No Dividend)



def black_scholes_dexter(S0,X,t,σ="",r=10,q=0.0,td=365):

# if(σ==""):σ =indiavix()

S0,X,σ,r,q,t = float(S0),float(X),float(σ/100),float(r/100),float(q/100),float(t/td)
#https://unofficed.com/black-scholes-model-options-calculator-google-sheet/

d1 = (math.log(S0/X)+(r-q+0.5*σ**2)*t)/(σ*math.sqrt(t))
#stackoverflow.com/questions/34258537/python-typeerror-unsupported-operand-types-for-float-and-int

#stackoverflow.com/questions/809362/how-to-calculate-cumulative-normal-distribution
Nd1 = (math.exp((-d1**2)/2))/math.sqrt(2*math.pi)
d2 = d1-σ*math.sqrt(t)
Nd2 = norm.cdf(d2)
call_theta =(-((S0*σ*math.exp(-q*t))/(2*math.sqrt(t))*(1/(math.sqrt(2*math.pi)))*math.exp(-(d1*d1)/2))-(r*X*math.exp(-r*t)*norm.cdf(d2))+(q*math.exp(-q*t)*S0*norm.cdf(d1)))/td
put_theta =(-((S0*σ*math.exp(-q*t))/(2*math.sqrt(t))*(1/(math.sqrt(2*math.pi)))*math.exp(-(d1*d1)/2))+(r*X*math.exp(-r*t)*norm.cdf(-d2))-(q*math.exp(-q*t)*S0*norm.cdf(-d1)))/td
call_premium =math.exp(-q*t)*S0*norm.cdf(d1)-X*math.exp(-r*t)*norm.cdf(d1-σ*math.sqrt(t))
put_premium =X*math.exp(-r*t)*norm.cdf(-d2)-math.exp(-q*t)*S0*norm.cdf(-d1)
call_delta =math.exp(-q*t)*norm.cdf(d1)
put_delta =math.exp(-q*t)*(norm.cdf(d1)-1)
gamma =(math.exp(-r*t)/(S0*σ*math.sqrt(t)))*(1/(math.sqrt(2*math.pi)))*math.exp(-(d1*d1)/2)
vega = ((1/100)*S0*math.exp(-r*t)*math.sqrt(t))*(1/(math.sqrt(2*math.pi))*math.exp(-(d1*d1)/2))
call_rho =(1/100)*X*t*math.exp(-r*t)*norm.cdf(d2)
put_rho =(-1/100)*X*t*math.exp(-r*t)*norm.cdf(-d2)

return call_theta,put_theta,call_premium,put_premium,call_delta,put_delta,gamma,vega,call_rho,put_rho