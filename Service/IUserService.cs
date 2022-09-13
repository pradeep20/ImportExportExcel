using SampleExcel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SampleExcel.Service
{
    public interface IUserService
    {
        List<User> GetUsers();

        bool SaveUsers(List<User> users);
    }
}
