using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class Dice : MonoBehaviour
{
    public List<Material> MaterialColors;
    public List<Texture> TexturesColors;

    const float x1_1 = -90f;
    const float x1_2 = 270f;
    const float xz2_1 = 0f;
    const float xz2_2 = 360f;
    const float xz2_3 = 180f;
    const float z3_1 = -90f;
    const float z3_2 = 270f;
    const float z4_1 = 90f;
    const float z4_2 = -270f;
    const float xz5_1 = 0;
    const float xz5_2 = 360f;
    const float x6_1 = 90f;
    const float x6_2 = -270f;

    const float fAccuracy = 10f;
    Transform t;
    Vector3 v3InitPos = Vector3.zero;

    Rigidbody rb;
    Transform tBowl;

    Renderer renderer;

    bool bSetColorOnce = false;

    void Start()
    {
        t = transform;
        //t.Rotate(new Vector3(Random.value * 360, Random.value * 360, Random.value * 360));
        v3InitPos = t.position;
        tBowl = GameObject.Find("Bowl").transform;
        rb = GetComponent<Rigidbody>();
        renderer = GetComponent<MeshRenderer>();
        /*
        Debug.Log(gameObject.name + " is " + GetValue() +
                ". Rotation = " + string.Format("({0},{1},{2})", t.eulerAngles.x, t.eulerAngles.y, t.eulerAngles.z));
                */
    }
    
    private void Update()
    {
        if (Input.GetKeyUp(KeyCode.A))
        {
            Debug.Log(gameObject.name + " is " + GetValue()+
                ". Rotation = "+string.Format("({0},{1},{2})",t.eulerAngles.x,t.eulerAngles.y,t.eulerAngles.z));
        }

        if (bSetColorOnce && !Rolling)
        {
            StartCoroutine(SetColor());
        }
    }
    /// <summary>
    /// Get the current value of the dice.
    /// </summary>
    /// <returns></returns>
    public int GetValue()
    {
        if (Rolling)
        {
            return -1;
        }
        int v;
        float x = t.eulerAngles.x;
        float z = t.eulerAngles.z;

        if (CheckEqual(x, x1_1,fAccuracy) ||CheckEqual(x,x1_2, fAccuracy))
            v = 1;
        else if (CheckEqual(x, xz2_1,xz2_2, fAccuracy) && CheckEqual(z, xz2_3, fAccuracy))
            v = 2;
        else if (CheckEqual(z, z3_1, fAccuracy) || CheckEqual(z, z3_2, fAccuracy))
            v = 3;
        else if (CheckEqual(z, z4_1, fAccuracy) || CheckEqual(z, z4_2, fAccuracy))
            v = 4;
        else if (CheckEqual(x, x6_1, fAccuracy) || CheckEqual(x, x6_2, fAccuracy))
            v = 6;
        else if (CheckEqual(x, xz5_1,xz5_2,10) && CheckEqual(z, xz5_1,xz5_2,10))
            v = 5;
        else
            v = 0;
 
        return v;
    }

    IEnumerator  SetColor()
    {
        //Debug.Log("SetColor now~_~");
        yield return null;
        if (renderer.material == null)
        {
            Debug.LogError(gameObject.name + " Dice material not found!");
            yield break ;
        }

        if (MaterialColors.Count != 6)
        {
            Debug.Log("Dice.cs MaterialColors's not set right,pls check the prefab");
            yield break;
        }

        if (TexturesColors.Count != 6)
        {
            Debug.Log("Dice.cs TexturesColors's not set right,pls check the prefab");
            yield break; 
        }

        int v = GetValue();
        if (v < 1 || v > 6)
        {
            //Debug.LogError("Dice.cs SetColor failed! Dice's value = " + v);
            StartCoroutine(BeginSetColor(0.1f));
            yield break;
            
        }
        //Debug.Log("v = " + v + ", MaterialColors = "+ MaterialColors[v-1].name);
        renderer.material = MaterialColors[v - 1];
        renderer.material.mainTexture = TexturesColors[v - 1];
        bSetColorOnce = false;
    }

    private bool CheckEqual(float a ,float b,float accuracy)
    {
        return Mathf.Abs(a - b) < accuracy;
    }

    private bool CheckEqual(float a, float b,float c, float accuracy)
    {
        return (Mathf.Abs(a - b) < accuracy || Mathf.Abs(a - c) < accuracy);
    }

    public bool Rolling
    {
        get
        {
            bool isRolling = true;
            if (rb && tBowl)
            {
                float d = Mathf.Abs(Vector3.Distance(t.position, tBowl.position));
                if (rb.velocity.sqrMagnitude < .01F && rb.angularVelocity.sqrMagnitude < .01F &&
                     d < .5f)
                {
                    isRolling = false;
                }
                else
                    isRolling = true;
            }
            return isRolling;
        }
    }

    public void DoRoll()
    {
        t.position = v3InitPos;
        t.Rotate(new Vector3(Random.value * 360, Random.value * 360, Random.value * 360));
        StartCoroutine(BeginSetColor());
    }

    IEnumerator BeginSetColor(float time = 0.1f)
    {
        yield return new WaitForSeconds(time);
        bSetColorOnce = true;
    }
}
