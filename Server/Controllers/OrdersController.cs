using Server.Classes;
using Server.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;


namespace Server.Controllers
{
    public class OrdersController : ApiController
    {
        OrdersContext db = new OrdersContext();

        public IEnumerable<Order> Get()
        {
            return db.Orders.ToList();
        }

        // GET api/<controller>/5
        public Order Get(int id)
        {
            return db.Orders.Find(id);
        }

        // POST api/<controller>
        public IHttpActionResult Post([FromBody]Order order)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }
            db.Orders.Add(order);
            db.SaveChanges();

            //DocCreator creator = new DocCreator();
            //creator.CreateDoc();
            return Ok();
        }

        // DELETE api/<controller>/5
        public IHttpActionResult Delete(int id)
        {
            Order order = db.Orders.Find(id);
            if (order != null)
            {
                db.Orders.Remove(order);
                db.SaveChanges();
                return Ok();
            }
            return NotFound();
        }
    }
}