using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.DotNet.MSIdentity.Shared;
using Microsoft.IdentityModel.Tokens;
using Spire.License;
using System.Collections;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;
using Vue.Models;

namespace it_api.Services
{
    public class AuthManager
    {
        private readonly UserManager<UserModel> _userManager;
        private readonly IConfiguration _configuration;
        private UserModel _user;
        public AuthManager(UserManager<UserModel> userManager,
            IConfiguration configuration)
        {
            _userManager = userManager;
            _configuration = configuration;
        }
        public async Task<bool> ValidateUser(string email, string password)
        {
            _user = await _userManager.FindByNameAsync(email);
            return (_user != null && await _userManager.CheckPasswordAsync(_user, password));
        }
        public async Task<string> CreateToken(UserModel user)
        {
            _user = user;
            var signingCredentials = GetSigningCredentials();
            var claim = await GetClaims();
            var tokenOptions = GenerateTokenOptions(signingCredentials, claim);
            return new JwtSecurityTokenHandler().WriteToken(tokenOptions);
        }
        public JwtSecurityToken decodeToken(string stream)
        {
            var handler = new JwtSecurityTokenHandler();
            var jsonToken = handler.ReadToken(stream);
            var tokenS = jsonToken as JwtSecurityToken;
            return tokenS;
        }
        public bool ValidateToken(string token, out JwtSecurityToken jwt)
        {
            var jwtSettings = _configuration.GetSection("JWT");
            var validationParameters = new TokenValidationParameters
            {
                ValidateIssuer = true,
                ValidIssuer = jwtSettings.GetSection("ValidIssuer").Value,
                ValidateAudience = true,
                ValidAudience = jwtSettings.GetSection("ValidAudience").Value,
                ValidateIssuerSigningKey = true,
                IssuerSigningKey = GetSecret(),
                ValidateLifetime = true
            };

            try
            {
                var tokenHandler = new JwtSecurityTokenHandler();
                tokenHandler.ValidateToken(token, validationParameters, out SecurityToken validatedToken);
                jwt = (JwtSecurityToken)validatedToken;

                return true;
            }
            catch (SecurityTokenValidationException ex)
            {
                jwt = null;
                // Log the reason why the token is not valid
                return false;
            }
        }
        private JwtSecurityToken GenerateTokenOptions(SigningCredentials signingCredentials, List<Claim> claims)
        {
            var jwtSettings = _configuration.GetSection("JWT");
            var expireValue = int.Parse(jwtSettings.GetSection("Expire").Value);
            var expireTime = DateTime.Now.AddDays(expireValue);
            var tokenOptions = new JwtSecurityToken(
                issuer: jwtSettings.GetSection("ValidIssuer").Value,
                audience: jwtSettings.GetSection("ValidAudience").Value,
                claims: claims,
                expires: expireTime,
                signingCredentials: signingCredentials
                );
            return tokenOptions;
        }
        private SigningCredentials GetSigningCredentials()
        {
            var secret = GetSecret();
            return new SigningCredentials(secret, SecurityAlgorithms.HmacSha256);
        }
        private SymmetricSecurityKey GetSecret()
        {
            var secretKey = _configuration.GetSection("JWT:Secret").Value;
            var keyByte = Encoding.UTF8.GetBytes(secretKey);
            var secret = new SymmetricSecurityKey(keyByte);
            return secret;
        }
        private async Task<List<Claim>> GetClaims()
        {
            var claims = new List<Claim> {
            new Claim(ClaimTypes.NameIdentifier, _user.Id),
            new Claim(ClaimTypes.Email, _user.Email),
            new Claim(ClaimTypes.Name, _user.UserName)
        };
            var roles = await _userManager.GetRolesAsync(_user);
            foreach (var role in roles)
            {
                claims.Add(new Claim(ClaimTypes.Role, role));
            }
            return claims;
        }

    }
}
