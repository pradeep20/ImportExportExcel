using SampleExcel.Context;
using SampleExcel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;

namespace SampleExcel.Service
{
    public class UserService : IUserService
    {
        DatabaseContext _dbContext = null;

        public UserService(DatabaseContext dbContext)
        {
            _dbContext = dbContext;
        }

        public List<User> GetUsers()
        {
            return _dbContext.Users.ToList();
        }

        public bool SaveUsers(List<User> users)
        {
            //foreach (var user in users)
            //{
            //    _dbContext.Add(user);
            //}

            _dbContext.Users.AddRange(users);
            _dbContext.SaveChanges();
            return true;
        }
    }
}
