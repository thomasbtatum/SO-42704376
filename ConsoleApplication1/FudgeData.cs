using System;
using System.Collections.Generic;
using System.Linq;

namespace ConsoleApplication1
{
    internal class FudgeData
    {
        public FudgeData()
        {
        }

        internal List<FudgeItem> Fudged()
        {
            return new[]
            {
                new FudgeItem() { Deaths=3,EventYear=2011, PSQReviewable=4},
                new FudgeItem() { Deaths=4,EventYear=2016, PSQReviewable=4},
                new FudgeItem() { Deaths=5,EventYear=2015, PSQReviewable=4},
                new FudgeItem() { Deaths=2,EventYear=2013, PSQReviewable=4},
                new FudgeItem() { Deaths=1,EventYear=2012, PSQReviewable=4}
            }.ToList();
        }
    }
}