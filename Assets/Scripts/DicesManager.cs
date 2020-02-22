using System;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.UI;

public class DicesManager : MonoBehaviour
{
    public Dropdown ddSelectDiceNum;
    public Button btnRollDice;

    public List<Dice> ListDices;
    
    void Start()
    {
        if (ddSelectDiceNum)
        {
            ddSelectDiceNum.onValueChanged.AddListener(ddSelectDiceNumOnValueChanged);
        }
        if (btnRollDice)
            btnRollDice.onClick.AddListener(btnRollDiceOnClicked);
    }

    private void ddSelectDiceNumOnValueChanged(int arg0)
    {
        ShowDice(ListDices.Count-arg0);
    }

    private void btnRollDiceOnClicked()
    {
        DoRoll();
    }

    private void ShowDice(int num)
    {
        if (num > ListDices.Count)
        {
            Debug.LogError("DicesManager.cs ShowDice num out of range!");
            return;
        }
        for(int i=0;i<ListDices.Count;i++)
        {
            ListDices[i].gameObject.SetActive(i < num);
        }
    }

    void Update()
    {
        if (Input.GetKeyUp(KeyCode.Space))
        {
            DoRoll();
        }
    }

    public void DoRoll()
    {
        foreach(var dice in ListDices)
        {
            if (dice.gameObject.activeInHierarchy)
                dice.DoRoll();
        }
    }
}
