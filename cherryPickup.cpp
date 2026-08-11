#include <bits/stdc++.h>
using namespace std;

class Solution
{
private:
    int find_max(vector<vector<int>> &grid, int i, int j1, int j2, int m, int n, vector<vector<vector<int>>> &dp)
    {

        if (j1 < 0 || j2 < 0 || j1 >= m || j2 >= m)
            return -1e9;
        if (dp[i][j1][j2] != -1)
            return dp[i][j1][j2];

        if (i == n - 1)
        {
            if (j1 == j2)
                return grid[i][j1];
            return grid[i][j1] + grid[i][j2];
        }

        int maxi = -1e9;

        for (int dj1 = -1; dj1 <= 1; dj1++)
        {
            for (int dj2 = -1; dj2 <= 1; dj2++)
            {

                int value;
                if (j1 == j2)
                {
                    value = grid[i][j1];
                    cout << "Single Pick -> " << " i-> " << i << " j1-> " << j1 << " " << value << endl;
                }

                else
                {
                    value = grid[i][j1] + grid[i][j2];
                }

                value += find_max(grid, i + 1, j1 + dj1, j2 + dj2, m, n, dp);

                maxi = max(maxi, value);
                dp[i][j1][j2] = maxi;
            }
        }

        return maxi;
    }

public:
    int cherryPickup(vector<vector<int>> &grid)
    {
        int n = grid.size();
        int m = grid[0].size();
        vector<vector<vector<int>>> dp(n, vector<vector<int>>(m, vector<int>(m, -1)));
        return find_max(grid, 0, 0, m - 1, m, n, dp);
        // int find_max(grid, i, j1, j2, m, n, dp);
    }
};

int main()
{
    Solution s1;
    int n, m;
    cin >> n >> m;
    vector<vector<int>> v(n, vector<int>(m));
    v = {{3, 1, 1}, {2, 5, 1}, {1, 5, 5}, {2, 1, 1}};
    cout << s1.cherryPickup(v) << endl;
}