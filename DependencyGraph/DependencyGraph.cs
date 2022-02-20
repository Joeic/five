// Skeleton implementation written by Joe Zachary for CS 3500, September 2013.
// Version 1.1 (Fixed error in comment for RemoveDependency.)
// Version 1.2 - Daniel Kopta 
//               (Clarified meaning of dependent and dependee.)
//               (Clarified names in solution/project structure.)

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace SpreadsheetUtilities
{

  /// <summary>
  /// (s1,t1) is an ordered pair of strings
  /// t1 depends on s1; s1 must be evaluated before t1
  /// 
  /// A DependencyGraph can be modeled as a set of ordered pairs of strings.  Two ordered pairs
  /// (s1,t1) and (s2,t2) are considered equal if and only if s1 equals s2 and t1 equals t2.
  /// Recall that sets never contain duplicates.  If an attempt is made to add an element to a 
  /// set, and the element is already in the set, the set remains unchanged.
  /// 
  /// Given a DependencyGraph DG:
  /// 
  ///    (1) If s is a string, the set of all strings t such that (s,t) is in DG is called dependents(s).
  ///        (The set of things that depend on s)    
  ///        
  ///    (2) If s is a string, the set of all strings t such that (t,s) is in DG is called dependees(s).
  ///        (The set of things that s depends on) 
  //
  // For example, suppose DG = {("a", "b"), ("a", "c"), ("b", "d"), ("d", "d")}
  //     dependents("a") = {"b", "c"}
  //     dependents("b") = {"d"}
  //     dependents("c") = {}
  //     dependents("d") = {"d"}
  //     dependees("a") = {}
  //     dependees("b") = {"a"}
  //     dependees("c") = {"a"}
  //     dependees("d") = {"b", "d"}
  /// </summary>
  public class DependencyGraph
  {
    /// <summary>
    /// Creates an empty DependencyGraph.
    /// </summary>
    public DependencyGraph()
    {
     
      graph = new Dictionary<String, HashSet<String>>();
      
      _size = 0;
    }
    // the structure of  the graph
    private Dictionary<String, HashSet<String>> graph;
    // the size of graph
    private int _size;




    /// <summary>
    /// The number of ordered pairs in the DependencyGraph.
    /// </summary>
    public int Size
    {
      get { return _size; }
    }



    /// <summary>
    /// The size of dependees(s).
    /// This property is an example of an indexer.  If dg is a DependencyGraph, you would
    /// invoke it like this:
    /// dg["a"]
    /// It should return the size of dependees("a")
    /// </summary>
    public int this[string s]
    {

      get
      {
        var c = 0;
        foreach (var t in graph)
        {
          if (t.Value.Contains(s))
            c++;
        }

        return c;
      }
    }


    /// <summary>
    /// Reports whether dependents(s) is non-empty.
    /// </summary>
    public bool HasDependents(string s)
    {
      //if no key return false
      if (graph.ContainsKey(s) == false)
        return false;
      else
      {
        return graph[s].Count > 0;
      }

    }


    /// <summary>
    /// Reports whether dependees(s) is non-empty.
    /// </summary>
    public bool HasDependees(string s)
    {
      var c = 0;
      foreach (var t in graph)
      {
        if (t.Value.Contains(s))
        {
          c++;
        }
      }

      return c == 0;
    }


    /// <summary>
    /// Enumerates dependents(s).
    /// </summary>
    public IEnumerable<string> GetDependents(string s)
    {

      if (graph.ContainsKey(s) == false)
        return new List<String>();
      else
      {
        return new HashSet<string>(graph[s]);
      }



    }

    /// <summary>
    /// Enumerates dependees(s).
    /// </summary>
    public IEnumerable<string> GetDependees(string s)
    {
      foreach (var t in graph)
      {
        if (t.Value.Contains(s))
        {
          yield return t.Key;
        }
      }
    }


    /// <summary>
    /// <para>Adds the ordered pair (s,t), if it doesn't exist</para>
    /// 
    /// <para>This should be thought of as:</para>   
    /// 
    ///   t depends on s
    ///
    /// </summary>
    /// <param name="s"> s must be evaluated first. T depends on S</param>
    /// <param name="t"> t cannot be evaluated until s is</param>        /// 
    public void AddDependency(string s, string t)
    {

      if (graph.ContainsKey(s))
      {
        // if s not in graph we should a new node
        if (graph[s].Contains(t) == false)
        {
          _size++;
         
          graph[s].Add(t);
        }
      }
      else
      {
        _size++;
        graph.Add(s, new HashSet<string>());
        graph[s].Add(t);
      }



    }


    /// <summary>
    /// Removes the ordered pair (s,t), if it exists
    /// </summary>
    /// <param name="s"></param>
    /// <param name="t"></param>
    public void RemoveDependency(string s, string t)
    {
      if (graph.ContainsKey(s))
      {
        if (graph[s].Contains(t))
        {
          graph[s].Remove(t);
          _size--;

        }
      }
    }


    /// <summary>
    /// Removes all existing ordered pairs of the form (s,r).  Then, for each
    /// t in newDependents, adds the ordered pair (s,t).
    /// </summary>
    public void ReplaceDependents(string s, IEnumerable<string> newDependents)
    {
      // if graph contain s,replace it 
      if (graph.ContainsKey(s))
      {
        _size -= graph[s].Count;
        graph[s].Clear();
        foreach (var t in newDependents)
        {
          graph[s].Add(t);
        }

        _size += graph[s].Count;
      }
      // if not contain , add a new node to graph
      else
      {
        graph.Add(s, new HashSet<string>(newDependents));
        _size += newDependents.Count<string>();
      }
    }


    /// <summary>
    /// Removes all existing ordered pairs of the form (r,s).  Then, for each 
    /// t in newDependees, adds the ordered pair (t,s).
    /// </summary>
    public void ReplaceDependees(string s, IEnumerable<string> newDependees)
    {
    
      foreach (var t in graph)
      {
        t.Value.Remove(s);
        _size--;
      }

      foreach (var t in newDependees)
      {

        if (graph.ContainsKey(t))
        {
          graph[t].Add(s);
          _size++;

        }
        else
        {
          graph.Add(t, new HashSet<string>());
          graph[t].Add(s);
          _size++;

        }
      }

    }

  }

}
